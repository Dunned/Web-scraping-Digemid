from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from openpyxl.workbook import Workbook
import time
import pandas as pd


class Automatizador :
    def empezarTarea(self,nombreCsv):
        #nombreCsv='codigosCSV.csv'
        csv=pd.read_csv(nombreCsv)
        codigos = csv['REGISTRO SANITARIO'].tolist()
        registrosNoRecuperados=[]
        df=pd.DataFrame(columns=[
            'REGISTRO SANITARIO',
            'NOMBRE',
            'FÓRMA FARMACÉUTICA',
            'TITULAR DEL REGISTRO',
            'GENÉRICO/MARCA',
            'TIPO DE RECETA',
            'ESTADO DEL REGISTRO',
            'FECHA VENCIMIENTO REGISTRO',
            'FABRICANTE',
            'PROCEDENCIA',
            'COD GRUPO TERAPEUTICO (COD ATC)',
            'DESCRIPCIÓN GRUPO TERAPEUTICO (COD ATC)',
            'VIDA ÚTI',
            'MOLÉCULA 1',
            'MOLÉCULA 2',
            'MOLÉCULA 3',
            'MOLÉCULA 4',
            'MOLÉCULA 5',
            'MOLÉCULA 6',
            'MOLÉCULAS ADICIONALES',
            'CONCENTRACIÓN 1',
            'CONCENTRACIÓN 2',
            'CONCENTRACIÓN 3',
            'CONCENTRACIÓN 4',
            'CONCENTRACIÓN 5',
            'CONCENTRACIÓN 6',
            'CONCENTRACIÓN ADICIONALES',
            'VIA DE ADMINISTRACIÓN',
            'DCI (DENOMINACIÓN COMÚN INTERNACIONAL)',
            'PRESENTACIÓN'])


        for cod in codigos:
            try:
                # Opciones de navegación
                options = webdriver.ChromeOptions()
                options.add_argument('--start-maximized')
                options.add_argument('--disable-extensions')

                driver_path = 'chromedriver_win32\\chromedriver.exe'

                driver = webdriver.Chrome(driver_path, chrome_options=options)

                driver.get('https://www.digemid.minsa.gob.pe/ProductosFarmaceuticos/principal/pages/Default.aspx')

                WebDriverWait(driver, 7) \
                    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                       'input#txtRegistro'))) \
                    .send_keys(cod)

                WebDriverWait(driver, 7) \
                    .until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                       'input#btnBuscar'))) \
                    .click()

                WebDriverWait(driver, 7) \
                    .until(EC.element_to_be_clickable((By.XPATH,
                                                       '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr')))

                nombre = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[3]').text
                formaFarmaceutica = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[4]').text
                titularRegistro = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[5]').text
                genericoMarca = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[6]').text
                tipoReceta = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[7]').text
                estado = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[8]').text

                WebDriverWait(driver, 7) \
                    .until(EC.element_to_be_clickable((By.XPATH,
                                                       '/html/body/form/div[3]/div[2]/div[3]/div/div/div[3]/div/div[2]/div/table/tbody/tr/td[1]'))) \
                    .click()

                WebDriverWait(driver, 7) \
                    .until(EC.element_to_be_clickable((By.XPATH,
                                                       '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div')))

                fechaVencimiento = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[1]/div[6]/b').text.lstrip()
                fabricante = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[6]/div[2]').text.lstrip()
                procedencia = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[7]/div[2]').text.lstrip()
                grupo = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[8]/div[2]').text
                grupo = grupo.split()
                codGrupoTerapeutico = grupo[0]
                grupo.pop(0)
                desGrupoTerapeutico = " ".join(grupo)
                del grupo
                vidaUtil = " "




                valirdarOrden = driver.find_element_by_xpath(
                    '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[10]/div').text.lstrip()
                #SIN SALTO LINEA
                if (valirdarOrden == 'COMPOSICION:'):
                    campoConcentracion = driver.find_element_by_xpath(
                        '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[11]/div').text
                    via = driver.find_element_by_xpath(
                        '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[13]').text.lstrip()
                    valirdarOrden2=driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[14]/div').text.lstrip()
                    if valirdarOrden2=='LIBERACION:':
                        presentacion=driver.find_element_by_xpath('/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[17]/div').text.lstrip()
                    else:
                        presentacion = driver.find_element_by_xpath(
                            '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[15]/div').text.lstrip()
                else:
                    campoConcentracion = driver.find_element_by_xpath(
                        '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[12]/div').text
                    via = driver.find_element_by_xpath(
                        '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[14]').text.lstrip()

                    valirdarOrden2 = driver.find_element_by_xpath(
                        '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[15]/div').text.lstrip()
                    if valirdarOrden2=='LIBERACION:':
                        presentacion = driver.find_element_by_xpath(
                            '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[18]/div').text.lstrip()

                    else:
                        presentacion = driver.find_element_by_xpath(
                            '/html/body/form/div[3]/div[2]/div[3]/div/div/div[1]/div/div/div/div[2]/div[16]/div').text.lstrip()



                campoConcentracion = campoConcentracion.split("\n")
                # LT,ML,G,MG,MCG,UI,U,%
                lt = []
                ml = []
                g = []
                mg = []
                mcg = []
                ui = []
                u = []
                p = []
                i = 0
                # print(campoConcentracion)
                for campo in campoConcentracion:
                    if i == 0:
                        i = i + 1
                        continue
                    campo = campo.lstrip()

                    campo2 = campo.lower()
                    campo2 = campo2.split()[-1]
                    # print(campo)
                    if campo2 == 'lt':
                        lt.append(campo)
                    if campo2 == 'ml':
                        ml.append(campo)
                    if campo2 == 'g':
                        g.append(campo)
                    if campo2 == 'mg':
                        mg.append(campo)
                    if campo2 == 'mcg':
                        mcg.append(campo)
                    if campo2 == 'ui':
                        ui.append(campo)
                    if campo2 == 'u':
                        mcg.append(campo)
                    if campo2 == '%':
                        mcg.append(campo)

                masMenosConcentraacion = [lt, ml, g, mg, mcg, ui, u, p]
                # print(mcg)
                moleculas = []  # 7
                concentraciones = []  # 7
                for categoria in masMenosConcentraacion:
                    if len(moleculas) < 6:
                        if categoria:
                            cateogiraN = []
                            categoriaV = []
                            medida = None
                            for valor in categoria:
                                valor = valor.split()
                                numero1 = valor.pop(-1)  # MG G UI
                                numero2 = valor.pop(-1)  # NUMERO
                                # VALOR SOLO QUEDA TEXTO
                                valor = " ".join(valor)
                                cateogiraN.append(valor)
                                categoriaV.append(float(numero2))
                                medida = numero1

                            for i in range(1, len(categoriaV)):
                                for j in range(0, len(categoriaV) - i):
                                    if (categoriaV[j + 1] > categoriaV[j]):
                                        aux = categoriaV[j]
                                        categoriaV[j] = categoriaV[j + 1]
                                        categoriaV[j + 1] = aux

                                        aux = cateogiraN[j]
                                        cateogiraN[j] = cateogiraN[j + 1]
                                        cateogiraN[j + 1] = aux

                            for i in range(0, len(cateogiraN)):
                                if len(moleculas) < 6:
                                    moleculas.append(cateogiraN[i])
                                    valorpushear = str(categoriaV[i]) + " " + medida
                                    concentraciones.append(valorpushear)
                                else:
                                    break
                    else:
                        moleculas.append('OTRAS MOLECULAS')
                        concentraciones.append('OTRAS CONCENTRACIONES')
                        break

                while len(moleculas) < 7:
                    moleculas.append(" ")
                    concentraciones.append(" ")

                df = df.append({
                    'REGISTRO SANITARIO': cod,
                    'NOMBRE': nombre,
                    'FÓRMA FARMACÉUTICA': formaFarmaceutica,
                    'TITULAR DEL REGISTRO': titularRegistro,
                    'GENÉRICO/MARCA': genericoMarca,
                    'TIPO DE RECETA': tipoReceta,
                    'ESTADO DEL REGISTRO': estado,
                    'FECHA VENCIMIENTO REGISTRO': fechaVencimiento,
                    'FABRICANTE': fabricante,
                    'PROCEDENCIA': procedencia,
                    'COD GRUPO TERAPEUTICO (COD ATC)': codGrupoTerapeutico,
                    'DESCRIPCIÓN GRUPO TERAPEUTICO (COD ATC)': desGrupoTerapeutico,
                    'VIDA ÚTI': " ",
                    'MOLÉCULA 1': moleculas[0],
                    'MOLÉCULA 2': moleculas[1],
                    'MOLÉCULA 3': moleculas[2],
                    'MOLÉCULA 4': moleculas[3],
                    'MOLÉCULA 5': moleculas[4],
                    'MOLÉCULA 6': moleculas[5],
                    'MOLÉCULAS ADICIONALES': moleculas[6],
                    'CONCENTRACIÓN 1': concentraciones[0],
                    'CONCENTRACIÓN 2': concentraciones[1],
                    'CONCENTRACIÓN 3': concentraciones[2],
                    'CONCENTRACIÓN 4': concentraciones[3],
                    'CONCENTRACIÓN 5': concentraciones[4],
                    'CONCENTRACIÓN 6': concentraciones[5],
                    'CONCENTRACIÓN ADICIONALES': concentraciones[6],
                    'VIA DE ADMINISTRACIÓN': via,
                    'DCI (DENOMINACIÓN COMÚN INTERNACIONAL)': " ",
                    'PRESENTACIÓN':presentacion
                }, ignore_index=True)

            except Exception as e:
                registrosNoRecuperados.append(cod)
                continue

        if(registrosNoRecuperados):
            df2 = pd.DataFrame(registrosNoRecuperados, columns=['Codigos No Encontrados'])
            df2.to_excel('noEncontrados.xlsx')


        df.to_excel('resultado.xlsx')
