import base64
import json
from datetime import datetime

import fitz
import cv2
import os
import shutil

import pyodbc
from PIL import Image
import easyocr
from funciones import Parser, borrar_calas

# url_origen = r"P:\Extraccion Calas - Jorge Castillo\Pruebas\Por Cargar"
url_origen = r"\\10.5.0.13\areas\ADMINISTRACION Y FINANZAS\LOGISTICA\90_RECURSOS\REPORTE DE CALA PARTICULARES\Bitacora Electronica\Por Cargar"

#url_origen = r"D:\OneDrive - Austral Group\Por Cargar"

def subir(url: str, planta: str):
    subidos = 0
    fallidos = 0
    error = ''
    ficheros = os.listdir(url)

    # Borrar las calas que serán subidas nuevamente
    borrar_calas(url)

    files = []
    for fichero in ficheros:
        if os.path.isfile(os.path.join(url, fichero)) and (fichero.endswith('.pdf') or fichero.endswith('.PDF')):
            files.append(fichero)

    print(files)

    for it in files:
        try:
            print("NOMBRE:", it)
            ress = []
            with fitz.open(os.path.join(url, it)) as doc:
                text = ""
                name = it[:-4]
                for page in doc:
                    text += page.get_text()

                try:
                    os.mkdir(url + "\\%s" % name)
                except OSError:
                    pass
                for i in range(len(doc)):
                    for img in doc.get_page_images(i):
                        xref = img[0]
                        pix = fitz.Pixmap(doc, xref)
                        if pix.n < 5:
                            pix.save(url + "\\%s\\p%s-%s.png" % (name, i, xref))
                        else:
                            pix1 = fitz.Pixmap(fitz.csRGB, pix)
                            pix1.save(url + "\\%s\\p%s-%s.png" % (name, i, xref))
                            pix1 = None
                        pix = None
                text = text.replace('�', '')

                ##################################################################

                ld = os.listdir(url + "\\%s" % name)
                imagenes = []
                for f in ld:
                    imagen = Image.open(url + "\\%s\\" % name + f)
                    if imagen.size == (1834, 615):
                        imagenes.append(imagen)

                for ic in imagenes:
                    icono = ic.convert("RGBA")

                    pixels = icono.load()
                    width, height = icono.size
                    for x in range(width):
                        for y in range(height):
                            r, g, b, a = pixels[x, y]
                            if y < 530 and x > 160 and (
                                    (r, g, b) != (128, 128, 128) or (r, g, b) in [(156, 224, 255), (224, 255, 255)]):
                                pixels[x, y] = (255, 255, 255, a)

                    icono.save(os.path.join(url, "imagen.png"))

                    image = cv2.imread(os.path.join(url, "imagen.png"))

                    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                    cv2.imwrite(os.path.join(url, "imagen.png"), image)

                    image = cv2.imread(os.path.join(url, "imagen.png"))
                    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                    canny = cv2.Canny(gray, 10, 150)

                    reader = easyocr.Reader(['es'], gpu=False)
                    result = reader.readtext(image, paragraph=False)
                    hpoint = result[0]

                    # posicion en y del intermedio del maximo valor de la izquierda
                    maximo_intermedio = (hpoint[0][3][1] + hpoint[0][0][1]) / 2

                    cnts, _ = cv2.findContours(canny, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

                    contador = 0
                    Ys = []
                    for c in cnts:
                        if c[0][0][0] > 160 and c[0][0][1] < 530:
                            contador += 1

                            epsilon = 0.01 * cv2.arcLength(c, True)
                            approx = cv2.approxPolyDP(c, epsilon, True)

                            x, y, w, h = cv2.boundingRect(approx)

                            # altura / diferencia de pixeles
                            valor = int(h * int(hpoint[1]) / (y + h - maximo_intermedio))
                            Ys.append([x, valor])  # ordenar por x
                            # altura 0 es "y+h"

                    Ys = sorted(Ys, key=lambda ys: ys[0])
                    Xs = result[-contador:]

                    for i in range(len(Xs)):
                        f = float(Xs[i][1])

                        if f >= 100:
                            f /= 10
                        # print(f)
                        Xs[i] = (Xs[i][0], f)

                    Xs = sorted(Xs, key=lambda xs: xs[1])

                    res = []
                    for i in range(len(Ys)):
                        res.append([Xs[i][1], Ys[i][1]])

                    # Pares x,y:
                    # for x in res:
                    #     print(x, end=" ")
                    # print()
                    ress.append(res)

                    # for (_, y) in Ys:
                    #     print(y, end=" ")
                    # print()
                    # for (_, number, _) in Xs:
                    #     print(float(number), end=" ")

                    # left-top right-top right-bot left-bot

            name = Parser(text, ress, planta)
            os.rename(os.path.join(url, it), os.path.join(url, name + ".pdf"))
            # url_des = r"P:\Extraccion Calas - Jorge Castillo\Pruebas\Por Cargar" + "\\" + planta
            url_des = r"\\10.5.0.13\areas\ADMINISTRACION Y FINANZAS\LOGISTICA\90_RECURSOS\REPORTE DE CALA PARTICULARES\Bitacora Electronica\Bitacoras Cargadas" + "\\" + planta
            os.path.join(url, name, ".pdf")
            shutil.move(url + "\\" + name + ".pdf", url_des + "\\" + name + ".pdf")
            imagen.close()
            for file in os.listdir(url + "\\" + str(it[:-4])):
                try:
                    os.remove(os.path.join(url, str(it[:-4]), file))
                except:
                    continue
            try:
                os.rmdir(os.path.join(url, str(it[:-4])))
            except:
                print(end="")
            subidos += 1

        except Exception as e:
            # PDF NO SE PUDO SUBIR
            print(e)
            print("ERROR: ", type(e).__name__, e.__cause__)
            try:
                os.stat(os.path.join(url, "Corregir"))
            except:
                os.mkdir(os.path.join(url, "Corregir"))

            shutil.move(os.path.join(url, it), os.path.join(url, "Corregir", it))
            fallidos += 1
            error = type(e).__name__

    try:
        os.remove(url + "\\" + "imagen.png")
    except:
        print(end="")

    return subidos, fallidos, error


cnxstr = "Driver={SQL Server};Server=10.5.0.52;Initial Catalog=POWERBI;Persist Security Info=False;User ID=UsrPortales;Password=$#ewo2001.2d;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=True;"

def final():
    # obtener pdfs de la db de correos
    cnxn = pyodbc.connect(cnxstr)
    cursor = cnxn.cursor()
    table = cursor.execute("select * from POWERBI.mprima.fact_faena_correos")
    indexes = []

    for raw in table.fetchall():
        indexes.append(raw[0])
        planta = ""
        if raw[3] == "COISHCO":
            planta = "COISHCO"
        elif raw[3] == "CHANCAY":
            planta = "CHANCAY"
        elif raw[3] == "PISCO":
            planta = "PISCO"
        elif raw[3] == "ILO":
            planta = "ILO"
        with open(os.path.join(url_origen, planta, raw[1]), "wb") as f:
            f.write(os.path.join(base64.b64decode(raw[2])))

    if indexes:
    # Hay elementos en indexes, por lo que construimos la consulta
        inject = "(" + ",".join(str(i) for i in indexes) + ")"
        sql_delete_query = "DELETE FROM POWERBI.mprima.fact_faena_correos WHERE Id IN " + inject
        cursor.execute(sql_delete_query)
        cursor.commit()

    else:
    # No hay elementos en indexes, por lo que no hay nada que borrar
        print("No hay registros para eliminar.")

    # inicio del proceso
    fecha_inicio = datetime.now().strftime('%d/%m/%Y %H:%M')

    plantas = ["Coishco", "Chancay", "Pisco", "Ilo"]
    subidos = 0
    fallidos = 0
    error = ''
    estado = ''
    for planta in plantas:
        borrar_calas(os.path.join(url_origen, planta))
        subidosi, fallidosi, errori = subir(os.path.join(url_origen, planta), planta)
        subidos += subidosi
        fallidos += fallidosi
        if errori != '':
            error += ',' + errori

    fecha_final = datetime.now().strftime('%d/%m/%Y %H:%M')
    if subidos != 0 and fallidos == 0:
        estado = 'EXITO'
    elif subidos == 0 and fallidos == 0:
        estado = 'VACIO'
    elif subidos != 0 and fallidos != 0:
        estado = 'SUBIDO CON ERRORES'
    else:
        estado = 'ERROR DE SUBIDA'

    cursor.execute('INSERT INTO POWERBI.mprima.TB_PROCESO_PDF (FECHOR_INI, FECHOR_FIN, ESTADO_PROCESO, NUMCANPDF, NUMCANPDF_ERR, OBSERVACIONES) VALUES (?,?,?,?,?,?)',
                   fecha_inicio, fecha_final, estado, subidos, fallidos, error)
    cursor.commit()
    cursor.close()

final()

