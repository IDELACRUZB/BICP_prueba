import time
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.service import Service
from twocaptcha import TwoCaptcha, NetworkException
from PIL import Image
from io import BytesIO
import pyautogui
import os
import glob
import shutil
import datetime
import random
import subprocess
import string
#para enviar email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class descargaReportes():
    def __init__(self):
        self.directoryPath = os.getcwd()
        self.defaultPathDownloads = self.directoryPath + r'\temp'
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("prefs", {
            # "download.default_directory": r"C:\Users\Usuario\Documents\terceriza\Robot\descargasPython\descargaRobotin",
            "download.default_directory": self.defaultPathDownloads,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        # Para Ignorar los errores de certificado SSL (La conexion no es privada)
        self.options.add_argument("--ignore-certificate-errors")
        self.pathDriver = "driver/chromedriver.exe"
        self.url = "https://10.95.224.27:9083/bicp/login.action"
        self.service = Service(self.pathDriver)
        # Configura el controlador de Chrome y abre la URL especificada
        #self.driver = webdriver.Chrome(service=self.service, options=self.options)
        self.driver = webdriver.Chrome(options=self.options)

    def reiniciar(self):
        self.__init__()

    # Funcion para obtener fechas
    def fecha(self, inicio = None, fin = None):
        if inicio is None:
            fechaActual = datetime.date.today()
            fechaSiguiente = fechaActual + datetime.timedelta(days=1)
        else:
            fechaActual = datetime.datetime.strptime(inicio, '%Y-%m-%d')
            if fin is None:
                fechaSiguiente = fechaActual + datetime.timedelta(days=1)
            else:
                fechaSiguiente = datetime.datetime.strptime(fin, '%Y-%m-%d')
            
        horaCero = datetime.time(0, 0, 0)
        fechaActualConHora = datetime.datetime.combine(fechaActual, horaCero)
        fechaSiguienteConHora = datetime.datetime.combine(fechaSiguiente, horaCero)
        formato = "%m/%d/%Y %H:%M:%S"
        formato2 = "%Y-%m-%d %H:%M:%S"
        
        faCH1 = fechaActualConHora.strftime(formato)
        fsCH1 = fechaSiguienteConHora.strftime(formato)
        faCH2 = fechaActualConHora.strftime(formato2)
        fsCH2 = fechaSiguienteConHora.strftime(formato2)
        faFormato1 = fechaActual.strftime("%m/%d/%Y")
        fsFormato1 = fechaSiguiente.strftime("%m/%d/%Y")
        fechaActual = fechaActual.strftime("%Y-%m-%d")
        fechaSiguiente = fechaSiguiente.strftime("%Y-%m-%d")
        fechas = {"hoyH":faCH1, 'mañanaH': fsCH1, 
                'hoyH2': faCH2, 'mañanaH2': fsCH2,
                'hoy': fechaActual, 'mañana': fechaSiguiente, 
                'hoyF1': faFormato1, 'mañanaF1': fsFormato1}
        return fechas

    #Funcion cuenta cantidad de archivos xlsx en carpeta de descarga
    def cantidadExcel(self):
        self.ruta_carpeta = self.defaultPathDownloads
        self.extension = '*.xlsx'
        self.patron_busqueda = os.path.join(self.ruta_carpeta, self.extension)
        self.archivos = glob.glob(self.patron_busqueda)
        self.cantidad_archivos = len(self.archivos)
        return self.cantidad_archivos

    def login(self):
        self.driver.get(self.url)
        time.sleep(1)

        # cierra ventana emergente
        self.elemento_cierre = self.driver.find_element(By.XPATH, '//*[@id="tipClose"]')
        self.elemento_cierre.click()
        time.sleep(1)
    
    # funcion para obtener el captcha
    def obtenerCaptcha(self):
        # Busca el elemento deseado utilizando un selector CSS
        self.element = self.driver.find_element(By.ID, 'validateimg')
        # Captura una imagen del elemento
        self.element.screenshot('capture.png')

        try:
            self.solver = TwoCaptcha('8f63da7191fe11e63148c3d8b28c71f2')
            self.id = self.solver.send(file='capture.png')
            time.sleep(10)
            self.codigoCaptcha = self.solver.get_result(self.id)
        except NetworkException:
            self.codigoCaptcha = 'error'

        return self.codigoCaptcha
    
    # funcion para inicio de sesion
    def inicioSesion(self, usuario, contrasena):
        # Busca los campos de usuario y contraseña y los llena con los valores especificados
        self.campoInputUser = self.driver.find_element(By.ID, 'username2')
        self.campoInputUser.clear()
        self.actions = ActionChains(self.driver)
        self.actions.move_to_element(self.campoInputUser)
        self.actions.click()
        self.actions.send_keys(usuario)
        self.actions.perform()
        time.sleep(1)

        self.campoInputPassword = self.driver.find_element(By.ID, 'password2')
        self.campoInputPassword.clear()
        self.actions = ActionChains(self.driver)
        self.actions.move_to_element(self.campoInputPassword)
        self.actions.click()
        self.actions.send_keys(contrasena)
        self.actions.perform()

        self.validacion_logueo = True
        while self.validacion_logueo:
            self.code = self.obtenerCaptcha()
            while self.code == 'error':
                self.refrescarCaptcha = self.driver.find_element(By.CLASS_NAME, 'validate_btn')
                self.refrescarCaptcha.click()
                self.code = self.obtenerCaptcha()
            else:
                pass
            self.code_field = self.driver.find_element(By.XPATH, '//*[@id="validate"]')
            self.code_field.send_keys(self.code)
            time.sleep(1)

            # presiona boton submit
            self.submit_button = self.driver.find_element(By.CLASS_NAME, 'login_submit_btn')
            self.submit_button.click()
            time.sleep(5)

            self.nueva_url = self.driver.current_url
            if self.nueva_url == self.url:
                self.validacion_logueo = True
                self.code_field.clear()
            else:
                self.validacion_logueo = False
        time.sleep(1)
    
    # Cierra la aplicacion
    def cerrarSesion(self):
        fun_logout = self.driver.find_element(By.ID, 'fun_logout')
        fun_logout.click()
        time.sleep(1)

        confirmaCierreSesion = self.driver.find_element(By.XPATH, '//*[@id="winmsg0"]/div/div[2]/div[3]/div[3]/span[1]/div/div')
        confirmaCierreSesion.click()
        time.sleep(3)

    # funcion para validar Sesion activa
    def validaSesionActiva(self):
        # Si la sesion esta activa
        try:
            self.continuar = self.driver.find_element(By.ID, 'usm_continue')
            self.sesionAbierta = True
        except NoSuchElementException:
            self.sesionAbierta = False

        if self.sesionAbierta:
            self.continuar.click()
        else:
            pass

    # funcion para validar cookies
    def validaSiExisteCookie(self):
        # para las cookies
        try:
            self.cookie = self.driver.find_element(By.CLASS_NAME, 'neterror')
            self.existeCookie = True
        except NoSuchElementException:
            self.existeCookie = False
        return self.existeCookie
    
    def gameOver(self):
        self.driver.quit()

    # +++++ 1. Funcion para descargar Reporte 1259 +++++
    def reporte1259(self, xpathBpo, xpathCamapana, fechaInicial=None, fechaFinal=None):
        # Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()
        time.sleep(1)

        # selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        # Dentro de la seccion de Resource Manager
        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID, 'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        # Dentro de la seccion Search Resource
        # Completa el TextBox con el codigo de la campaña
        codigoCampana = 1259
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID, 'searchCondtion_input_value')
        textBox.clear()
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="1259"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        # Seccion de Input de parámetros
        self.driver.switch_to.frame('view8adf609b6f1e1f27016f21e66d51024c_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH']
        fechaFin = self.fecha(inicio, fin)['mañanaH']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === BPO ===
        btnBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnBPO.click()
        time.sleep(2)

        # Desmarca la seleccion por default
        checkBoxDefault = self.driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
        checkBoxDefault.click()
        time.sleep(1)
        # Marca el checbox de la campaña desseada -- 1er argumento funcion
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBpo)
        checkBoxCampana.click()
        time.sleep(1)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)
        # === * ===
        # === Campaign ===
        if xpathCamapana != None:
            btnCampaign = self.driver.find_element(By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
            btnCampaign.click()
            time.sleep(2)

            # Desmarca la seleccion por default
            checkBoxDefault = self.driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
            checkBoxDefault.click()
            time.sleep(1)
            # Marca el checbox de la campaña desseada -- 1er argumento funcion
            checkBoxCampana = self.driver.find_element(By.XPATH, xpathCamapana)
            checkBoxCampana.click()
            time.sleep(1)

            BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
            BotonOk.click()
            time.sleep(1)
        else:
            pass
        # ==== * ===

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        self.cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
            
        #Descargar
        iconoDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        iconoDowloand.click()
        time.sleep(1)
        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        self.cantidadExcelFinal = self.cantidadExcelinicial
        while self.cantidadExcelFinal == self.cantidadExcelinicial:
            time.sleep(1)
            self.cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin Funcion Reporte 1259 +++++

    # +++++ 2. Funcion para descargar Reporte 401 +++++
    def reporte401(self, xpathBPO, txtCampana, xpathCampana, xpathAgentWorkgroup=None, fechaInicial=None, fechaFinal=None):
        # Menu Principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()
        time.sleep(1)

        # selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        # Dentro del Resource Manager
        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID, 'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        # Dentro de la seccion Search Resource
        # Completa el TextBox con el codigo de la campaña
        codigoCampana = 401
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID, 'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="401"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        # Seccion de Input de parámetros
        self.driver.switch_to.frame('view8adf609c6387afda01638b7f47b90570_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio,fin)['hoyH']
        fechaFin = self.fecha(inicio,fin)['mañanaH']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === BPO ===
        btnBPO = self.driver.find_element(
            By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnBPO.click()
        time.sleep(2)

        # desmarca selecion Default All
        checkBoxDefault = self.driver.find_element(
            By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
        checkBoxDefault.click()
        time.sleep(1)

        # Selecciona BPO
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxCampana.click()
        time.sleep(1)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        # === Campaign ===
        btnCampana = self.driver.find_element(By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
        btnCampana.click()
        time.sleep(2)

        # desmarca selecion Default All
        checkBoxDefault = self.driver.find_element(By.CSS_SELECTOR, 'input[type="checkbox"][value="9999"]')
        checkBoxDefault.click()
        time.sleep(1)

        # ingresar texto busqueda en textbox
        texto = txtCampana
        textBox = self.driver.find_element(By.NAME, 'startwith')
        textBox.send_keys(texto)
        time.sleep(1)

        botonSearch = self.driver.find_element(By.NAME, 'search')
        botonSearch.click()
        time.sleep(2)

        # Selecciona Campaign
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathCampana)
        checkBoxCampana.click()
        time.sleep(1)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        # === Agent Workgroup ===
        agentWkg = xpathAgentWorkgroup
        if agentWkg != None:
            btnAgentWorkgroup = self.driver.find_element(By.XPATH, '//*[@id="Agent WorkgroupViewControl"]/table/tbody/tr/td[5]/input')
            btnAgentWorkgroup.click()
            time.sleep(2)

            checkBox1034 = self.driver.find_element(By.XPATH, agentWkg)
            checkBox1034.click()
            time.sleep(1)

            BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
            BotonOk.click()
            time.sleep(1)
        else:
            pass

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

        # Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass
        
        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin Funcion Reporte 401 +++++

    # +++++ 3. Funcion para descargar Reporte 112 +++++
    def reporte112(self, xpathBPO, fechaInicial=None, fechaFinal=None):
        # Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        # selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID, 'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 112
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID, 'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="112"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input parametros
        self.driver.switch_to.frame('view8adf609c54f300b50154f46383e501a8_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoy']
        fechaFin = self.fecha(inicio, fin)['mañana']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === BPO ===
        # Seleccionar campaña
        
        btnCampana = self.driver.find_element(
            By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnCampana.click()
        time.sleep(2)

        # desmarca toda la selecion por Default SelectAll
        checkBoxselectAll = self.driver.find_element(
            By.CSS_SELECTOR, 'input[type="checkbox"][value="selectAll"]')
        checkBoxselectAll.click()
        checkBoxselectAll.click()
        time.sleep(1)

        # Selecciona BPO
        checkBoxBPO = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxBPO.click()
        time.sleep(1)

        BotonOk = self.driver.find_element(
            By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

        # Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(
            By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin funcion reporte 112 +++++

    # +++++ 4. Funcion Reporte 1261 +++++
    def reporte1261(self, xpathBPO, xpathActivity, xpatkCampana, textoC, fechaInicial=None, fechaFinal=None):
        # Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()
        time.sleep(1)
        # selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)
        # Dentro del Resource Manager
        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID, 'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 1261
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID, 'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="1261"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        # *** MOV 1 ***
        # === Fecha D - 1 ===
        self.driver.switch_to.frame('view8adf609b6f1e1f27016f21e66fd10255_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyF1']
        fechaFin = self.fecha(inicio, fin)['mañanaF1']
        # Formato mes/dia/anio
        fechaD_1 = fechaInicio

        textBoxFecha = self.driver.find_element(By.ID, 'rpt_param_Value0')
        textBoxFecha.clear()
        textBoxFecha.send_keys(fechaD_1)
        time.sleep(1)

        # === BPO ===
        btnBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnBPO.click()
        time.sleep(2)

        # Selecciona BPO
        checkBoxBPO = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxBPO.click()
        time.sleep(2)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(2)

        # === Activity ===
        btnActivity = self.driver.find_element(By.XPATH, '//*[@id="ActivityViewControl"]/table/tbody/tr/td[5]/input')
        btnActivity.click()
        time.sleep(2)

        texto = textoC
        textBox = self.driver.find_element(By.NAME, 'startwith')
        textBox.send_keys(texto)
        time.sleep(2)

        botonSearch = self.driver.find_element(By.NAME, 'search')
        botonSearch.click()
        time.sleep(2)

        # Selecciona Activity
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathActivity)
        checkBoxCampana.click()
        time.sleep(5)

        BotonOk = self.driver.find_element(
            By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(2)

        # === Campaign ===
        btnCampaign = self.driver.find_element(
            By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
        btnCampaign.click()
        time.sleep(2)

        texto = textoC
        textBox = self.driver.find_element(By.NAME, 'startwith')
        textBox.send_keys(texto)
        time.sleep(2)

        botonSearch = self.driver.find_element(By.NAME, 'search')
        botonSearch.click()
        time.sleep(2)

        # Selecciona Campaign
        checkBoxCampana = self.driver.find_element(By.XPATH, xpatkCampana)
        checkBoxCampana.click()
        time.sleep(2)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(2)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

        # Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass 
        
        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin Reporte 1261 +++++

    # +++++ 5. Funcion Reporte 90 +++++
    def reporte90(self, xpathCampana, xpathBPO, xpathAgentGroup, fechaInicial=None, fechaFinal=None):
        # Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()
        time.sleep(1)
        # selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)
        # Dentro del Resource Manager
        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID, 'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 90
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID, 'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="90"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        # Dentro del Input source
        self.driver.switch_to.frame('view8adf649b5237e8b7015352f0feb700d9_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH']
        fechaFin = self.fecha(inicio, fin)['mañanaH']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === Campaign ===
        btnCampaign = self.driver.find_element(
            By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
        btnCampaign.click()
        time.sleep(2)

        # Selecciona Campaign
        seleccionDefault = self.driver.find_element(
            By.XPATH, '/html/body/div[13]/ul/li[1]/span')
        seleccionDefault.click()
        time.sleep(1)
        seleccionDefault.click()
        time.sleep(1)

        checkBoxCampana = self.driver.find_element(By.XPATH, xpathCampana)
        checkBoxCampana.click()
        time.sleep(2)

        btnCampaign.click()
        time.sleep(2)

        # === BPO ===
        btnBPO = self.driver.find_element(
            By.XPATH, '//*[@id="rpt_param_ComboxId3"]/div/div')
        btnBPO.click()
        time.sleep(2)

        # Selecciona BPO
        seleccionDefault = self.driver.find_element(
            By.XPATH, '/html/body/div[14]/ul/li[1]/span')
        seleccionDefault.click()
        time.sleep(1)

        checkBoxBPO = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxBPO.click()
        time.sleep(2)

        btnBPO.click()
        time.sleep(2)

        # === Agent Group ===
        btnActivity = self.driver.find_element(
            By.XPATH, '//*[@id="AgentWorkGroupViewControl"]/table/tbody/tr/td[5]/input')
        btnActivity.click()
        time.sleep(2)

        # Selecciona Agent Group
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathAgentGroup)
        checkBoxCampana.click()
        time.sleep(5)

        BotonOk = self.driver.find_element(
            By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(2)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

        # Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(
            By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin Reporte 90 +++++

    # +++++ 6. Funcion Reporte 43 +++++
    def reporte43(self, xpathBPO, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 43
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="43"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input resource
        self.driver.switch_to.frame('view8adf609c511ba97601511ba994360060_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH']
        fechaFin = self.fecha(inicio, fin)['mañanaH']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)
        # === BPO ====
        comboBox = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
        comboBox.click()
        time.sleep(2)

        #desmarca selecion Default All
        checkBoxDefault = self.driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
        checkBoxDefault.click()
        time.sleep(1)

        #Selecciona BPO
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxCampana.click()
        time.sleep(1)

        comboBox.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))

        #Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass
        
        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin funcion Reporte 43 +++++

    # +++++ 7. Funcion Reporte 26 +++++
    def reporte26(self, xpathCampana, xpathBPO, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 26
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="26"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input resource
        self.driver.switch_to.frame('view8adf609c511ba97601511ba991d20056_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyF1']
        fechaFin = self.fecha(inicio, fin)['mañanaF1']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === Campaign Name ===
        comboBoxCN = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
        comboBoxCN.click()
        time.sleep(2)

        #desmarca selecion Default All
        checkBoxDefault = self.driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
        checkBoxDefault.click()
        time.sleep(1)

        #Selecciona Campaign Name
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathCampana)
        checkBoxCampana.click()
        time.sleep(1)

        comboBoxCN.click()
        time.sleep(2)

        # === BPO ====
        comboBoxBPO = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId3"]/div/div')
        comboBoxBPO.click()
        time.sleep(2)

        #desmarca selecion Default All
        checkBoxDefault = self.driver.find_element(By.XPATH, '/html/body/div[14]/ul/li[1]/span')
        checkBoxDefault.click()
        time.sleep(1)

        #Selecciona BPO
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxCampana.click()
        time.sleep(1)

        comboBoxBPO.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))    

        #Descarga
        btnDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        btnDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)

        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass
        
        self.driver.refresh()
        time.sleep(1)   
    # +++++ Fin funcon Reporte 26 +++++

    # +++++ 8. Funcion Reporte 194 +++++
    def reporte194(self, xpathBPO, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 194
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="194"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input resource
        self.driver.switch_to.frame('view8adf609c5becac34015bf11be608063a_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH2']
        fechaFin = self.fecha(inicio, fin)['mañanaH2']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === BPO ====
        comboBoxBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[3]/div/span/button/span[1]')
        comboBoxBPO.click()
        time.sleep(2)

        #Selecciona BPO
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxCampana.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
        
        #Descarga
        iconoDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        iconoDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)
        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin funcon Reporte 194 +++++

    # +++++ 9. Funcion Reporte 192 +++++
    def reporte192(self, xpathBPO, xpathAgentGroup, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 192
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="192"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input resource
        self.driver.switch_to.frame('view8adf609c5becac34015bf11be4bf0631_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH2']
        fechaFin = self.fecha(inicio, fin)['mañanaH2']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === BPO ====
        comboBoxBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[3]/div/span/button/span[1]')
        comboBoxBPO.click()
        time.sleep(2)

        #Selecciona BPO
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxCampana.click()
        time.sleep(1)

        # === Agent Group ===
        comboAgentGroup = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId4"]/div/div')
        comboAgentGroup.click()
        time.sleep(2)

        CheckboxDefault = self.driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
        CheckboxDefault.click()
        time.sleep(1)

        checkboxAgentGroup = self.driver.find_element(By.XPATH, xpathAgentGroup)
        checkboxAgentGroup.click()
        time.sleep(1)

        comboAgentGroup.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
        
        #Descarga
        iconoDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        iconoDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)
        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin funcon Reporte 194 +++++

    # +++++ 10. Funcion Reporte 418 +++++
    def reporte418(self, xpathCampana, xpathBPO, xpathAgentGroup, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()
        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)
        #Dentro del Resource Manager
        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 418
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="418"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        # Dentro del Input source
        self.driver.switch_to.frame('view8adf609c64556ca701645a17ef88026a_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyH']
        fechaFin = self.fecha(inicio, fin)['mañanaH']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === Campaign ===
        btnCampaign = self.driver.find_element(By.XPATH, '//*[@id="CampaignViewControl"]/table/tbody/tr/td[5]/input')
        btnCampaign.click()
        time.sleep(2)

        #Selecciona Campaign
        checkBoxCampana = self.driver.find_element(By.XPATH, xpathCampana)
        checkBoxCampana.click()
        time.sleep(1)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        # === BPO ===
        btnBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnBPO.click()
        time.sleep(2)

        #Selecciona BPO
        checkBoxBPO = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxBPO.click()
        time.sleep(2)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        # === Agent Group ===
        btnActivity = self.driver.find_element(By.XPATH, '//*[@id="Agent Group NameViewControl"]/table/tbody/tr/td[5]/input')
        btnActivity.click()
        time.sleep(2)

        #Selecciona Agent Group
        checkBoxAgentGroup = self.driver.find_element(By.XPATH, xpathAgentGroup)
        checkBoxAgentGroup.click()
        time.sleep(2)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(2)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(10)

        #Descarga
        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible tiempo maximo 300seg = 5 min
        WebDriverWait(self.driver, 300).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
        
        #Descarga
        iconoDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        iconoDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)
        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)
    # +++++ Fin Reporte 418 +++++

    # +++++11. Funcion Reporte 1392 +++++
    def reporte1392(self, xpathCampaign, xpathBPO, fechaInicial=None, fechaFinal=None):
        #Menu principal
        systemMenu = self.driver.find_element(By.ID, 'systemMenu')
        systemMenu.click()

        time.sleep(1)
        #selecciona Resource manager del menu principal
        resourceManager = self.driver.find_element(By.ID, 's_800')
        resourceManager.click()
        time.sleep(5)

        self.driver.switch_to.frame('tabPage_800_iframe')
        self.driver.switch_to.frame('container')
        lupa = self.driver.find_element(By.ID,'searchResource')
        lupa.click()
        time.sleep(1)
        self.driver.switch_to.default_content()
        self.driver.switch_to.default_content()

        codigoCampana = 1392
        self.driver.switch_to.frame('searchResource_iframe')
        textBox = self.driver.find_element(By.ID,'searchCondtion_input_value')
        textBox.send_keys(codigoCampana)
        time.sleep(1)

        searchBtn = self.driver.find_element(By.ID, 'searchBtn')
        searchBtn.click()
        time.sleep(3)

        reporte = self.driver.find_element(By.XPATH, '//a[@title="1392"]')
        reporte.click()
        self.driver.switch_to.default_content()
        time.sleep(5)

        #dentro del input resource
        self.driver.switch_to.frame('view8adf609c74632c1b0174671eb753033c_iframe')

        #Periodo
        inicio = fechaInicial
        fin = fechaFinal
        fechaInicio = self.fecha(inicio, fin)['hoyF1']
        fechaFin = self.fecha(inicio, fin)['mañanaF1']

        startTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value0"]')
        startTime.clear()
        startTime.send_keys(fechaInicio)
        time.sleep(1)

        endTime = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_Value1"]')
        endTime.clear()
        endTime.send_keys(fechaFin)
        time.sleep(1)

        # === campaign ====
        comboBoxCampana = self.driver.find_element(By.XPATH, '//*[@id="rpt_param_ComboxId2"]/div/div')
        comboBoxCampana.click()
        time.sleep(2)

        #Selecciona Campaign    
        checkBoxDefault = self.driver.find_element(By.XPATH, '/html/body/div[13]/ul/li[1]/span')
        checkBoxDefault.click()
        time.sleep(1)

        checkBoxCampana = self.driver.find_element(By.XPATH, xpathCampaign)
        checkBoxCampana.click()
        time.sleep(1)

        comboBoxCampana.click()
        time.sleep(1)

        # === BPO ===
        btnBPO = self.driver.find_element(By.XPATH, '//*[@id="BPOViewControl"]/table/tbody/tr/td[5]/input')
        btnBPO.click()
        time.sleep(2)

        #Selecciona BPO
        checkBoxBPO = self.driver.find_element(By.XPATH, xpathBPO)
        checkBoxBPO.click()
        time.sleep(2)

        BotonOk = self.driver.find_element(By.XPATH, '//*[contains(@id, "btnOk_rpt_param_")]/div/div')
        BotonOk.click()
        time.sleep(1)

        BotonOk2 = self.driver.find_element(By.ID, 'rpt_param_OkBtn')
        BotonOk2.click()
        time.sleep(1)

        cantidadExcelinicial = self.cantidadExcel()
        # Esperar hasta que la imagen loading ya no esté visible
        WebDriverWait(self.driver, 60).until_not(EC.visibility_of_element_located((By.ID, "loadingImg")))
        
        #Descarga
        iconoDowloand = self.driver.find_element(By.CLASS_NAME, 'ico_download')
        iconoDowloand.click()
        time.sleep(1)

        descargarExcel = self.driver.find_element(By.ID, 'downFullExcelMenuItemLiId')
        descargarExcel.click()
        time.sleep(1)

        excel2013 = self.driver.find_element(By.ID, 'downExcel2007Id')
        excel2013.click()
        time.sleep(1)

        confirmacionFinal = self.driver.find_element(By.ID, 'btnOk_downExcel2003Or2007WinId')
        confirmacionFinal.click()
        #time.sleep(15)
        #Valida que la descarga concluya
        cantidadExcelFinal = cantidadExcelinicial
        while cantidadExcelFinal == cantidadExcelinicial:
            time.sleep(1)
            cantidadExcelFinal = self.cantidadExcel()
        else:
            pass

        self.driver.refresh()
        time.sleep(1)

    # +++++ Fin funcon Reporte 1392 +++++

    # Funcion que reubicará las descargas en sus respectivas carpetas
    def renombrarReubicar(self, nuevoNombre, carpetaDestino):
        self.ruta_descargas = self.directoryPath + r'/temp'
        self.archivos_descargados = sorted(glob.glob(os.path.join(self.ruta_descargas, '*')), key=os.path.getmtime, reverse=True)
        # Comprobar si hay archivos descargados
        if len(self.archivos_descargados) > 0:
            self.ultimo_archivo = self.archivos_descargados[0]
            nuevo_nombre = f'{nuevoNombre}.xlsx'
            carpeta_destino = carpetaDestino
            # Comprobar si la carpeta de destino existe, si no, crearla
            if not os.path.exists(carpeta_destino):
                os.makedirs(carpeta_destino)
            # Ruta completa del archivo de destino
            self.ruta_destino = os.path.join(carpeta_destino, nuevo_nombre)
            # Mover el archivo a la carpeta de destino con el nuevo nombre
            shutil.move(self.ultimo_archivo, self.ruta_destino)


    # Funcion que crea el nombre del reporte
    def nombreReporte(self, name, fechaD0 = True):
        inicio='2023-10-20'
        if fechaD0:
            fechaHora = datetime.datetime.now()
            fecha = fechaHora.strftime("%Y%m%d_%H%M%S")
            aleatorio = str(random.randint(100, 999))
            nameFile = name + fecha + '_' + aleatorio
        else:
            h = datetime.datetime.now()
            hora = h.strftime('%H%M%S')
            fechan = datetime.datetime.strptime(inicio, '%Y-%m-%d')
            fechan = fechan + datetime.timedelta(days=1)
            fecha = fechan.strftime("%Y%m%d_")
            aleatorio = str(random.randint(100, 999))
            nameFile = name + fecha + hora + '_' + aleatorio
            
        return nameFile


