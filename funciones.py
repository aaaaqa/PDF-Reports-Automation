import fitz
import pyodbc
# from sqlalchemy import create_engine
from datetime import datetime
import os


class tabla_resumen:
    def __init__(self):
        self.DES_COD_UNICO = str()
        self.DES_NUMBER = str()
        self.DES_NOMBR_COMUN = str()
        self.DES_NOMBR_CIENT = str()
        self.NUM_PESCA_DECLA = "0.0"
        self.NUM_COMPO_CAPTU = "0.0"
        self.NUM_PORCE_JUVEN = "0.0"
        self.NUM_TOTAL_CAPTU = "0.0"


def isNumeric(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def parse_tabla_resumen(txt, idx, CodigoUnico, ListaFilasResumen):
    # print("N°\t|\tNomb Com\t|\tNombre Cientifico\t|\tPesca Decl.\t|\tComp.\t|\t% Juveniles\t|\tTotalCapt")

    while txt[idx] != 'Juveniles':
        idx += 1
    idx += 1
    while txt[idx] != 'Total':
        obj = tabla_resumen()
        obj.DES_COD_UNICO = CodigoUnico

        # Especie
        obj.DES_NUMBER = txt[idx]
        idx += 1
        # Nombre Comun
        obj.DES_NOMBR_COMUN = txt[idx]
        while txt[idx + 1].isupper() and txt[idx + 1] != 'N/E':
            obj.DES_NOMBR_COMUN += " " + txt[idx + 1]
            idx += 1
        idx += 1
        # Nombre Cientifico
        if not isNumeric(txt[idx][0]):
            obj.DES_NOMBR_CIENT = txt[idx]
            idx += 1
            while not isNumeric(txt[idx][0]):
                obj.DES_NOMBR_CIENT += " " + txt[idx]
                idx += 1
        else:
            obj.DES_NOMBR_CIENT = ""
        obj.NUM_PESCA_DECLA = txt[idx]
        if isNumeric(txt[idx + 1][0]) and isNumeric(txt[idx + 2][0]) and (isNumeric(txt[idx + 3][0]) and not isNumeric(txt[idx + 4][0]) or txt[idx + 3] == 'Total'):
            idx += 1
            obj.NUM_COMPO_CAPTU = txt[idx]
            idx += 1
            obj.NUM_PORCE_JUVEN = txt[idx]
        elif isNumeric(txt[idx + 1][0]) and (
                isNumeric(txt[idx + 2][0]) and not isNumeric(txt[idx + 3][0]) or txt[idx + 2] == 'Total'):
            idx += 1
            obj.NUM_COMPO_CAPTU = txt[idx]
            obj.NUM_PORCE_JUVEN = "0.0"
        else:
            obj.NUM_COMPO_CAPTU = "0.0"
            obj.NUM_PORCE_JUVEN = "0.0"
            idx += 1
        idx += 1

        obj.NUM_PESCA_DECLA = obj.NUM_PESCA_DECLA.replace(",", "")
        # print(obj.DES_NUMBER, obj.DES_NOMBR_COMUN, obj.DES_NOMBR_CIENT, obj.NUM_PESCA_DECLA, obj.NUM_COMPO_CAPTU, obj.NUM_PORCE_JUVEN,
        #       "Total", sep="\t|\t")
        # LISTO
        ListaFilasResumen.append(obj)
    TotalCapturado = txt[idx + 3]
    TotalCapturado = TotalCapturado.replace(",", "")
    idx += 4
    return ListaFilasResumen, TotalCapturado, idx


def Parser(txt, ListaGraficas, planta):
    # Cnx SQL
    cnxstr = "Driver={SQL Server};Server=10.5.0.52;Initial Catalog=POWERBI;Persist Security Info=False;User ID=UsrPortales;Password=$#ewo2001.2d;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=True;"

    cnxn = pyodbc.connect(cnxstr)

    cursor = cnxn.cursor()

    # TABLA PRINCIPAL
    # ID = str()
    CodigoUnico = str()
    FechaEmision = str()
    Embarcacion = str()
    Armador = str()
    PermisoPesca = str()
    Matricula = str()
    CapBodegaAuto = str()
    FechaRegistro = str()  # date
    FechaInicio = str()  # date
    OrigenInicio = str()
    PuertoZarpe = str()
    FechaFin = str()  # date
    OrigenFin = str()
    NCalas = int()

    # RESUMEN DE CAPTURAS DE FAENA
    NEspecie = str()
    NombreComun = str()
    NombreCientifico = str()
    PescaDeclarada = float()
    ComposicionCaptura = float()
    PorcentajeJuveniles = float()
    TotalCapturado = float()

    # DETALLE CALA
    NombreCala = str()
    FechaRegistroCala = str()
    OrigenCala = str()
    EspecieObjetivo = str()
    FechaInicioCala = str()
    LatitudInicial = str()
    LongitudInicial = str()
    DistCalaAnterior = str()
    CapturaT = str()
    FechaFinCala = str()
    LatitudFinal = str()
    LongitudFinal = str()
    Total = str()

    # DETALLE CALA SUBTABLA
    NCala = str()
    NombreComunCala = str()
    NombreCientificoCala = str()
    ComposicionCapturaCala = str()
    NEjemplaresJuvenilesCala = str()
    NindividuosMuestraCala = str()
    PorcentajeJuvenilesCala = str()

    # FechaCarga = datetime.now().strftime('%d/%m/%Y %H:%M')
    FechaCarga = datetime.now()

    txt = txt.replace('\n', ' ')
    txt = txt.split(' ')
    txt = [i for i in txt if i != '']
    # print(txt)

    CambiarFechaInicial = True
    CambiarFechaFinal = True

    subir_resumen = True

    ListaCalasGrafica = []

    for idx in range(len(txt)):
        palabra = txt[idx]
        # print(palabra)

        # Para las gráficas:
        if idx+2 < len(txt) and txt[idx] + txt[idx+2] == "CALA-":
            nombreCala = txt[idx] + " " + txt[idx+1] + " " + txt[idx+2] + " " + txt[idx+3]
            ListaCalasGrafica.append(nombreCala)
        # Fin para las gráficas

        if palabra == 'Embarcación':
            Embarcacion += txt[idx + 2]
            i = 3
            while txt[idx + i] != 'Matrícula':
                Embarcacion += " " + txt[idx + i]
                i += 1
        elif palabra == 'Matrícula':
            Matricula += txt[idx + 2]
        elif palabra == 'Armador':
            Armador += txt[idx + 5]
            i = 6
            while txt[idx + i] != 'Cap.':
                Armador += " " + txt[idx + i]
                i += 1
        elif palabra == 'Cap.':
            CapBodegaAuto = txt[idx + 4]
        elif palabra == 'Permiso':
            if txt[idx + 4] != 'DETALLE':
                print(end="")
            else:
                PermisoPesca += ""
        elif idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "DETALLEDEFAENA":
            CodigoUnico += txt[idx + 4]

            sFechaRegistro = txt[idx + 9] + " " + txt[idx + 10]
            FechaRegistro = datetime.strptime(sFechaRegistro, '%d/%m/%Y %H:%M')
            idx += 10
        elif idx + 1 < len(txt) and txt[idx] + txt[idx + 1] == "FechaInicio" and CambiarFechaInicial:
            sFechaInicio = txt[idx + 3] + " " + txt[idx + 4]
            FechaInicio = datetime.strptime(sFechaInicio, '%d/%m/%Y %H:%M')
            CambiarFechaInicial = False
        elif idx + 1 < len(txt) and txt[idx] + txt[idx + 1] == "FechaFin" and CambiarFechaFinal:
            sFechaFin = txt[idx + 3] + " " + txt[idx + 4]
            FechaFin = datetime.strptime(sFechaFin, '%d/%m/%Y %H:%M')
            CambiarFechaFinal = False
        elif idx + 1 < len(txt) and txt[idx] + txt[idx + 1] == "OrigenInicio":
            OrigenInicio += txt[idx + 3]
            i = 4
            while txt[idx + i] != 'OrigenFin':
                OrigenInicio += " " + txt[idx + i]
                i += 1
        elif txt[idx] == 'OrigenFin':
            OrigenFin += txt[idx + 2]
            i = 3
            while txt[idx + i] + txt[idx + i + 1] + txt[idx + i + 2] != "PuertodeZarpe":
                OrigenFin += " " + txt[idx + i]
                i += 1
        elif idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "PuertodeZarpe":
            PuertoZarpe += txt[idx + 4]
            i = 5
            while txt[idx + i] + txt[idx + i + 1] + txt[idx + i + 2] != "NúmerodeCalas":
                PuertoZarpe += " " + txt[idx + i]
                i += 1
        elif idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "NúmerodeCalas":
            NCalas = int(txt[idx + 4])
        elif idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "FechadeEmisión":
            sFechaEmision = txt[idx + 3] + " " + txt[idx + 4]
            FechaEmision = datetime.strptime(sFechaEmision, '%d/%m/%Y %H:%M')
        else:
            ############################################## RESUMEN DE CAPTURAS DE LA FAENA
            if idx + 5 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] + txt[idx + 3] + txt[idx + 4] + txt[idx + 5] == "RESUMENDECAPTURASDELAFAENA" and subir_resumen:
                ListaFilasResumen = [] # va a ser necesario para almacenar la captura total antes de insertar las filas de los resumenes de cala
                ##########
                # codigo #
                ##########
                subir_resumen = False
                ListaFilasResumen, TotalCapturado, idx = parse_tabla_resumen(txt, idx, CodigoUnico, ListaFilasResumen)

                print("\tCodigo Unico\t|\tDES_NUMBER\t|\tDES_NOMB_COMUN\t|\tDES_NOMB_CIENT\t|\tPESCA DECLARADA\t|\tCOMPO CAPT\t|\t%JUVENILES\t|\tTOTAL")
                for o in ListaFilasResumen:
                    print(CodigoUnico, o.DES_NUMBER, o.DES_NOMBR_COMUN, o.DES_NOMBR_CIENT, float(o.NUM_PESCA_DECLA), float(o.NUM_COMPO_CAPTU), float(o.NUM_PORCE_JUVEN), float(TotalCapturado), sep="\t|\t")
                    cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_RESUMEN (DES_CODIG_UNICO, DES_NUMER, DES_NOMBR_COMUN, DES_NOMBR_CIENT, NUM_PESCA_DECLA, NUM_COMPO_CAPTU, NUM_PORCE_JUVEN, NUM_TOTAL_CAPTU) VALUES (?,?,?,?,?,?,?,?)",
                                   CodigoUnico, o.DES_NUMBER, o.DES_NOMBR_COMUN, o.DES_NOMBR_CIENT, float(o.NUM_PESCA_DECLA), float(o.NUM_COMPO_CAPTU), float(o.NUM_PORCE_JUVEN), float(TotalCapturado))  # El none se cambiará cuando tenga el total capturado
                # print()

            ############################################## DETALLES DE CALAS UNITARIAS:
            if idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "DETALLEDECALA":
                # parsear cada cala
                # NombreCala = "Cala " + txt[idx+3]
                ID_DETALLE_CALA = txt[idx+3]
                NombreCala = txt[idx] + " " + txt[idx + 1] + " " + txt[idx + 2] + " " + txt[idx + 3]
                idx += 4
                if txt[idx] + txt[idx + 1] + txt[idx + 2] == "FechadeRegistro":
                    sFechaRegistroCala = txt[idx + 4] + " " + txt[idx + 5]
                    FechaRegistroCala = datetime.strptime(sFechaRegistroCala, '%d/%m/%Y %H:%M')
                idx += 6
                while txt[idx] != "Origen":
                    idx += 1

                OrigenCala = txt[idx + 2]
                idx += 2
                while txt[idx + 1] + txt[idx + 2] + txt[idx + 3] + txt[idx + 4] != "DistanciaRespectoCalaAnterior":
                    OrigenCala += " " + txt[idx + 1]
                    idx += 1
                DistCalaAnterior = txt[idx + 7]
                idx += 7
                while idx + 1 < len(txt) and txt[idx] + txt[idx + 1] != "EspecieObjetivo":
                    idx += 1
                idx += 3
                if idx + 1 < len(txt) and txt[idx] + txt[idx + 1] == "CapturaTotal":
                    EspecieObjetivo = ""
                else:
                    EspecieObjetivo = txt[idx]

                while not isNumeric(txt[idx][0]):
                    idx += 1
                CapturaT = txt[idx]
                idx += 1
                while idx + 1 < len(txt) and txt[idx] + txt[idx + 1] != "FechaInicio":
                    idx += 1
                idx += 3
                sFechaInicioCala = txt[idx] + " " + txt[idx + 1]
                FechaInicioCala = datetime.strptime(sFechaInicioCala, '%d/%m/%Y %H:%M')
                while idx + 1 < len(txt) and txt[idx] + txt[idx + 1] != "FechaFin":
                    idx += 1
                idx += 3
                sFechaFinCala = txt[idx] + " " + txt[idx + 1]
                FechaFinCala = datetime.strptime(sFechaFinCala, '%d/%m/%Y %H:%M')
                while txt[idx] != 'Latitud':
                    idx += 1
                idx += 2
                LatitudInicial = txt[idx] + " " + txt[idx + 1] + " " + txt[idx + 2]
                while txt[idx] != 'Latitud':
                    idx += 1
                idx += 2
                LatitudFinal = txt[idx] + " " + txt[idx + 1] + " " + txt[idx + 2]
                while txt[idx] != 'Longitud':
                    idx += 1
                idx += 2
                LongitudInicial = txt[idx] + " " + txt[idx + 1] + " " + txt[idx + 2]
                while txt[idx] != 'Longitud':
                    idx += 1
                idx += 2
                LongitudFinal = txt[idx] + " " + txt[idx + 1] + " " + txt[idx + 2]

                DistCalaAnterior = DistCalaAnterior.replace(",", "")
                CapturaT = CapturaT.replace(",", "")

                # FACT_FAENA_TERCEROS_CALAS *********************************
                print("Nombre cada Cala\t|\tFechaRegistroCala\t|\tOrigenCala\t|\tEspecieObj\t|\t--FechaInicioCala--\t|\tLatitudInicial\t|\tLongitudInicial\t|\tDistAnt\t|\tCapT\t|\tFechaFinCala\t|\tLatitudFinal\t|\tLongitudFinal\t|\tTotal")
                print(NombreCala, FechaRegistroCala, OrigenCala, EspecieObjetivo, FechaInicioCala, LatitudInicial, LongitudInicial, DistCalaAnterior, CapturaT, FechaFinCala, LatitudFinal, LongitudFinal, Total, sep="\t|\t")
                cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_CALAS (DES_CODIG_UNICO, DES_NUMER, DES_NOMBR_CALAS, FEC_REGIS, DES_ORIGE, DES_ESPEC_OBJEC, FEC_INICI, NUM_LATIT_INICI, NUM_LONGI_INICI, NUM_DISTA_CALAA, NUM_CAPTU_TOTAL, FEC_FINAL, NUM_LATIT_FINAL, NUM_LONGI_FINAL, NUM_TOTAL) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                               CodigoUnico, int(NombreCala[-1]), NombreCala, FechaRegistroCala, OrigenCala, EspecieObjetivo, FechaInicioCala, LatitudInicial, LongitudInicial, float(DistCalaAnterior), float(CapturaT), FechaFinCala, LatitudFinal, LongitudFinal, None)

                ############################################## DETALLE CALA SUBTABLA
                if EspecieObjetivo != "":
                    # Cala no vacía
                    # ListaCalasGrafica.append(NombreCala)
                    while txt[idx] + txt[idx + 1] != '%Juveniles':
                        idx += 1
                    idx += 2

                    print("N°\t|\tNomb Com\t|\tNombre Cientifico\t|\tComposicion\t|\tEjemp\t|\tIndiv\t|\t%Juveniles")

                    while txt[idx] != 'Total' and txt[idx] != 'REPORTE':
                        # Al finalizar el bucle deberían insertarse la cantidad de subtablas de la cala "n"
                        # Especie
                        NCala = txt[idx]
                        idx += 1
                        # Nombre Comun
                        NombreComunCala = txt[idx]
                        while txt[idx + 1].isupper() and txt[idx + 1] != 'N/E':
                            NombreComunCala += " " + txt[idx + 1]
                            idx += 1
                        idx += 1
                        # Nombre Cientifico
                        if not isNumeric(txt[idx][0]):
                            NombreCientificoCala = txt[idx]
                            idx += 1
                            while not isNumeric(txt[idx][0]):
                                NombreCientificoCala += " " + txt[idx]
                                idx += 1
                        else:
                            NombreCientificoCala = ""
                        ComposicionCapturaCala = txt[idx]
                        if isNumeric(txt[idx+1][0]) and isNumeric(txt[idx+2][0]) and isNumeric(txt[idx+3][0]) and (isNumeric(txt[idx+4][0]) and not isNumeric(txt[idx+5][0]) or txt[idx+4] in ['Total', 'RESUMEN']):
                            idx += 1
                            NEjemplaresJuvenilesCala = txt[idx]

                            idx += 1
                            NindividuosMuestraCala = txt[idx]

                            idx += 1
                            PorcentajeJuvenilesCala = txt[idx]
                        elif isNumeric(txt[idx+1][0]) and isNumeric(txt[idx+2][0]) and (isNumeric(txt[idx+3][0]) and not isNumeric(txt[idx+4][0]) or txt[idx+3] in ['Total', 'RESUMEN']):
                            idx += 1
                            NEjemplaresJuvenilesCala = txt[idx]

                            idx += 1
                            NindividuosMuestraCala = txt[idx]

                            PorcentajeJuvenilesCala = "0.0"
                        elif isNumeric(txt[idx+1][0]) and (isNumeric(txt[idx+2][0]) and not isNumeric(txt[idx+3][0]) or txt[idx+4] in ['Total', 'RESUMEN']):
                            idx += 1
                            NEjemplaresJuvenilesCala = txt[idx]
                            NindividuosMuestraCala = "0"
                            PorcentajeJuvenilesCala = "0.0"
                        else:
                            NEjemplaresJuvenilesCala = "0"
                            NindividuosMuestraCala = "0"
                            PorcentajeJuvenilesCala = "0.0"

                        print(NCala, NombreComunCala, NombreCientificoCala, ComposicionCapturaCala, NEjemplaresJuvenilesCala, NindividuosMuestraCala, PorcentajeJuvenilesCala, sep="\t|\t")
                        cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_CALAS_DETALLE (DES_CODIG_UNICO, IDD_DETAL_CALAS, DES_NOMBR_CALAS, DES_NUMER, DES_NOMBR_COMUN, DES_NOMBR_CIENT, NUM_COMPO_CAPTU, NUM_EJEMP_JUVEN, NUM_INDIV_MUEST, NUM_PORCE_JUVEN) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                       CodigoUnico, ID_DETALLE_CALA, NombreCala, NCala, NombreComunCala, NombreCientificoCala, float(ComposicionCapturaCala), int(NEjemplaresJuvenilesCala), int(NindividuosMuestraCala), float(PorcentajeJuvenilesCala))

                        idx += 1
                        if idx + 5 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] + txt[idx + 3] + txt[idx + 4] + txt[idx + 5] == "RESUMENDECAPTURASDELAFAENA" and subir_resumen:
                            ListaFilasResumen = []  # va a ser necesario para almacenar la captura total antes de insertar las filas de los resumenes de cala
                            ListaFilasResumen, TotalCapturado, idx = parse_tabla_resumen(txt, idx, CodigoUnico, ListaFilasResumen)
                            idx += 12
                            subir_resumen = False
                            for o in ListaFilasResumen:
                                print(CodigoUnico, o.DES_NUMBER, o.DES_NOMBR_COMUN, o.DES_NOMBR_CIENT, float(o.NUM_PESCA_DECLA), float(o.NUM_COMPO_CAPTU), float(o.NUM_PORCE_JUVEN), float(TotalCapturado), sep="\t|\t")
                                cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_RESUMEN (DES_CODIG_UNICO, DES_NUMER, DES_NOMBR_COMUN, DES_NOMBR_CIENT, NUM_PESCA_DECLA, NUM_COMPO_CAPTU, NUM_PORCE_JUVEN, NUM_TOTAL_CAPTU) VALUES (?,?,?,?,?,?,?,?)",
                                               CodigoUnico, o.DES_NUMBER, o.DES_NOMBR_COMUN, o.DES_NOMBR_CIENT, float(o.NUM_PESCA_DECLA), float(o.NUM_COMPO_CAPTU), float(o.NUM_PORCE_JUVEN), float(TotalCapturado))  # El none se cambiará cuando tenga el total capturado
                        # limpiar
                        NCala = str()
                        NombreComunCala = str()
                        NombreCientificoCala = str()
                        ComposicionCapturaCala = '0.0'
                        NEjemplaresJuvenilesCala = '0'
                        NindividuosMuestraCala = '0'
                        PorcentajeJuvenilesCala = '0.0'

                else:
                    # print("N°\t|\tNomb Com\t|\tNombre Cientifico\t|\tComposicion\t|\tEjemp\t|\tIndiv\t|\t%Juveniles")

                    NombreComunCala = str()
                    NombreCientificoCala = str()
                    ComposicionCapturaCala = '0.0'
                    NEjemplaresJuvenilesCala = '0'
                    NindividuosMuestraCala = '0'
                    PorcentajeJuvenilesCala = '0.0'

                    # SUBIR CALAS VACIAS
                    while txt[idx] + txt[idx + 1] != '%Juveniles':
                        idx += 1
                    idx += 2
                    NCala = txt[idx]
                    # print(NCala, NombreComunCala, NombreCientificoCala, ComposicionCapturaCala,
                    #       NEjemplaresJuvenilesCala, NindividuosMuestraCala, PorcentajeJuvenilesCala, sep="\t|\t")
                    cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_CALAS_DETALLE (DES_CODIG_UNICO, IDD_DETAL_CALAS, DES_NOMBR_CALAS, DES_NUMER, DES_NOMBR_COMUN, DES_NOMBR_CIENT, NUM_COMPO_CAPTU, NUM_EJEMP_JUVEN, NUM_INDIV_MUEST, NUM_PORCE_JUVEN) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                   CodigoUnico, ID_DETALLE_CALA, NombreCala, NCala, NombreComunCala, NombreCientificoCala, float(ComposicionCapturaCala), int(NEjemplaresJuvenilesCala), int(NindividuosMuestraCala), float(PorcentajeJuvenilesCala))
                print()

        # NO TOCAR!
        idx += 1

    print("\tCODIGO UNICO\t|\t---FECHAEMISION---\t|\tBARCO\t|\t--------ARMADOR--------\t|\tPERMISO PESCA\t|\tMATRICULA\t|\tCAP BOD\t|\t--FECHA REGISTRO--\t|\t--FECHA INICIO--\t|\t--ORIGEN--\t|\tPUE\t|\t----FECHA FIN----\t|\tORIGEN FIN\t|\tNCALAS")
    print(CodigoUnico, FechaEmision, Embarcacion, Armador, PermisoPesca, Matricula, CapBodegaAuto, FechaRegistro,
          FechaInicio, OrigenInicio, PuertoZarpe, FechaFin, OrigenFin, NCalas, sep="\t|\t")
    # LISTO
    cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS (DES_CODIG_UNICO, FEC_EMISI, DES_EMBAR, DES_RAZON_ARMAD, DES_PERMI_PESCA, DES_MATRI, NUM_CAPAC_BODEG, FEC_REGIS, FEC_INICI, DES_ORIGE_INICI, DES_PUERT_ZARPE, FEC_FINAL, DES_ORIGE_FINAL, NUM_CALAS, DES_RUTAS_ARCHI, FEC_CARGA) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   CodigoUnico, FechaEmision, Embarcacion, Armador, PermisoPesca, Matricula, CapBodegaAuto, FechaRegistro, FechaInicio, OrigenInicio, PuertoZarpe, FechaFin, OrigenFin, NCalas, None, FechaCarga)

    print(ListaCalasGrafica)
    print(ListaGraficas)

    index = 0

    # ¡¡¡------------DESCOMENTAR-------------!!!
    for grafica in ListaGraficas:
        # [Xs, Ys]
        for e in grafica:
            cursor.execute("INSERT INTO POWERBI.mprima.FACT_FAENA_TERCEROS_GRAFICO (FECHA_CARGA, DES_CODIG_UNICO, IDD_DETAL_CALAS ,DES_NOMBR_CALAS, NUM_CALAS, NUM_EJEX, NUM_EJEY) VALUES (?,?,?,?,?,?,?)",
                            FechaCarga, CodigoUnico, int(ListaCalasGrafica[index][5]), ListaCalasGrafica[index], None, float(e[0]), float(e[1]))
        index += 1

    # COD_20220112_HM
    cursor.commit()
    cursor.close()

    FechaEmision = FechaEmision.strftime("%d/%m/%Y %H:%M:%S")
    # print(FechaEmision)
    Embarcacion = Embarcacion.replace(" ", "_")
    nombre = "RFC_" + CodigoUnico + "_" + FechaEmision[:2] + "_" + FechaEmision[3:5] + "_" + FechaEmision[6:10] + "_" + FechaEmision[11:13] + "_" + FechaEmision[14:16] + "_" + Embarcacion + "_" + planta
    # print(FechaCarga)
    print(nombre)
    return nombre


def borrar_calas(url):
    cnxstr = "Driver={SQL Server};Server=10.5.0.52;Initial Catalog=POWERBI;Persist Security Info=False;User ID=UsrPortales;Password=$#ewo2001.2d;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=True;"

    cnxn = pyodbc.connect(cnxstr)
    cursor = cnxn.cursor()
    ficheros = os.listdir(url)
    files = []
    for fichero in ficheros:
        # print(fichero)
        if os.path.isfile(os.path.join(url, fichero)) and fichero.endswith('.pdf'):
            files.append(fichero)
    # print(files)
    for it in files:
        CodigoUnico = ""
        try:
            with fitz.open(os.path.join(url, it)) as doc:
                txt = ""
                for page in doc:
                    txt += page.get_text()
                txt = txt.replace('�', '')
                txt = txt.replace('\n', ' ')
                txt = txt.split(' ')
                txt = [i for i in txt if i != '']

                for idx in range(len(txt)):
                    if idx + 2 < len(txt) and txt[idx] + txt[idx + 1] + txt[idx + 2] == "DETALLEDEFAENA":
                        CodigoUnico += txt[idx + 4]
                        break
                # print(CodigoUnico)
                cursor.execute("DELETE from POWERBI.mprima.FACT_FAENA_TERCEROS where DES_CODIG_UNICO = ?", str(CodigoUnico))
                cursor.execute("DELETE from POWERBI.mprima.FACT_FAENA_TERCEROS_CALAS where DES_CODIG_UNICO = ?", str(CodigoUnico))
                cursor.execute("DELETE from POWERBI.mprima.FACT_FAENA_TERCEROS_RESUMEN where DES_CODIG_UNICO = ?", str(CodigoUnico))
                cursor.execute("DELETE from POWERBI.mprima.FACT_FAENA_TERCEROS_CALAS_DETALLE where DES_CODIG_UNICO = ? ", str(CodigoUnico))
                cursor.execute("DELETE from POWERBI.mprima.FACT_FAENA_TERCEROS_GRAFICO where DES_CODIG_UNICO = ?", str(CodigoUnico))
        except:
            continue

    cursor.commit()
    cursor.close()

