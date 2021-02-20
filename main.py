from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementNotSelectableException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import csv
import sys


options = Options()
options.headless = True
#options.page_load_strategy = 'normal'
profile = webdriver.FirefoxProfile()
profile.set_preference("network.http.pipelining", True)
profile.set_preference("network.http.proxy.pipelining", True)
profile.set_preference("network.http.pipelining.maxrequests", 8)
profile.set_preference("content.notify.interval", 500000)
profile.set_preference("content.notify.ontimer", True)
profile.set_preference("content.switch.threshold", 250000)
profile.set_preference("browser.cache.memory.capacity", 65536) # Increase the cache capacity.
profile.set_preference("browser.startup.homepage", "about:blank")
profile.set_preference("reader.parse-on-load.enabled", False) # Disable reader, we won't need that.
profile.set_preference("browser.pocket.enabled", False) # Duck pocket too!
profile.set_preference("loop.enabled", False)
profile.set_preference("browser.chrome.toolbar_style", 1) # Text on Toolbar instead of icons
profile.set_preference("browser.display.show_image_placeholders", False) # Don't show thumbnails on not loaded images.
profile.set_preference("browser.display.use_document_colors", False) # Don't show document colors.
profile.set_preference("browser.display.use_document_fonts", 0) # Don't load document fonts.
profile.set_preference("browser.display.use_system_colors", True) # Use system colors.
profile.set_preference("browser.formfill.enable", False) # Autofill on forms disabled.
profile.set_preference("browser.helperApps.deleteTempFileOnExit", True) # Delete temprorary files.
profile.set_preference("browser.shell.checkDefaultBrowser", False)
profile.set_preference("browser.startup.homepage", "about:blank")
profile.set_preference("browser.startup.page", 0) # blank
profile.set_preference("browser.tabs.forceHide", True) # Disable tabs, We won't need that.
profile.set_preference("browser.urlbar.autoFill", False) # Disable autofill on URL bar.
profile.set_preference("browser.urlbar.autocomplete.enabled", False) # Disable autocomplete on URL bar.
profile.set_preference("browser.urlbar.showPopup", False) # Disable list of URLs when typing on URL bar.
profile.set_preference("browser.urlbar.showSearch", False) # Disable search bar.
profile.set_preference("extensions.checkCompatibility", False) # Addon update disabled
profile.set_preference("extensions.checkUpdateSecurity", False)
profile.set_preference("extensions.update.autoUpdateEnabled", False)
profile.set_preference("extensions.update.enabled", False)
profile.set_preference("general.startup.browser", False)
profile.set_preference("plugin.default_plugin_disabled", False)
profile.set_preference("permissions.default.image", 2) # Image load disabled again

driver = webdriver.Firefox(options=options,executable_path='D:\\chromedriver\geckodriver.exe')
driver.implicitly_wait(10)
options.page_load_strategy = 'eager'
#driver = webdriver.Firefox(executable_path='D:\\chromedriver\geckodriver.exe')


driver.get("https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/login.jsf")

username = driver.find_element_by_id("frmLogin:txtLogin")
username.clear()
username.send_keys("{login}")
print ("username colocado")
password = driver.find_element_by_name("frmLogin:j_id8")
password.clear()
password.send_keys("{senha}")
print ("password colocado")

driver.find_element_by_name("frmLogin:j_id22").click()



items = []

with open('nota_fiscal_servico_2021-02-01_2021-02-11.csv', encoding="cp1252") as csvfile:
    dict = {}
    csvReader = csv.reader(csvfile)
    i =0
    j = 0
    for row in csvReader:
        items.append(row)
        dict["form1:cpfCnpjTomador"] = items[i][j]
        dict["form1:razaoSocialTomador"] = items[i][j+1]
        dict["form1:cepTomador"] = items[i][j+2]
        dict["form1:enderecoTomador"] = items[i][j+3]
        dict["form1:bairroTomador"] = items[i][j+4]
        dict["form1:cmbUfTomador"] = items[i][j+5]
        dict["form1:cmbMunicipioTomador"] = items[i][j+6]
        dict["form1:valorUnitarioItem"] = items[i][j + 7]

        try:

            #driver.implicitly_wait(10)
            driver.get("https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/bemVindo.jsf")
            print ("login efetivado")
            driver.refresh()
            time.sleep(1)
            driver.refresh()
            driver.find_element_by_name("j_id4:j_id27:3:j_id30").click()
            time.sleep(1)
            driver.get("https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/emissaoNFSe/emissaoNFSeInsercao.jsf")
            print ("Passo 1 - Formulário Inicial")
            time.sleep(1)

            cpfTomador = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"form1:cpfCnpjTomador")))
            cpfTomador.send_keys(dict["form1:cpfCnpjTomador"])
            razaoSocial = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"form1:razaoSocialTomador")))
            razaoSocial.send_keys(dict["form1:razaoSocialTomador"])
            print ("CPF e Razão Social Inseridos")
            time.sleep(1)
            endTomador = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.ID,"form1:enderecoTomador")))
            endTomador.clear()
            endTomador.send_keys(dict["form1:enderecoTomador"])
            print ("endereco inserido")
            cepTomador = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:cepTomador")))
            cepTomador.clear()
            cepTomador.send_keys(dict["form1:cepTomador"])
            print ("cep inserido")

            bairroTomador = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:bairroTomador")))
            bairroTomador.clear()
            bairroTomador.send_keys(dict["form1:bairroTomador"])
            print ("bairro inserido")

            wait2 = WebDriverWait (driver,10)

            #selectUF = Select(driver.find_element_by_name('form1:cmbUfTomador')) #select UF = MA
            #selectUF.select_by_value(dict["form1:cmbUfTomador"])
            element_UF = wait2.until(EC.element_to_be_selected(driver.find_element_by_id("form1:cmbUfTomador")))
            select_uf = Select(element_UF)
            select_uf.select_by_visible_text("MA")
            print ("UF inserida")
            #time.sleep (1)
            #selectMunicipio = Select(driver.find_element_by_name('form1:cmbMunicipioTomador')) #select Municipio
            #selectMunicipio.select_by_value(dict["form1:cmbMunicipioTomador"]) #value da opção
            element_Mun = wait2.until(EC.element_to_be_selected(driver.find_element_by_id("form1:cmbUfTomador")))
            select_Mun = Select(element_Mun)
            select_Mun.select_by_value(dict["form1:cmbMunicipioTomador"])
            print ("Municipio Inserido")
            # time.sleep (1)

            #WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "form1:j_id313"))).click()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "form1:j_id313"))).click()
            #driver.find_element_by_name('form1:j_id313').click() #avança para a segunda parte do formulário
            print ("Passo 2 - Página Atividades")
            time.sleep(1)

            #el = driver.find_element_by_id('form1:cmbAtividades')
            el = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:cmbAtividades")))
            for option in el.find_elements_by_tag_name('option'):
                if option.text == '773909900 - ALUGUEL DE OUTRAS MAQUINAS E EQUIPAMENTOS COMERCIAIS E INDUSTRIAIS NAO ESPECIFICADOS ANTERIORMENTE, SEM OPERADOR':
                    option.click() # select() in earlier versions of webdriver
                    break
            time.sleep(1)

            selectServ = Select(driver.find_element_by_name('form1:j_id381'))
            selectServ.select_by_value('0304')


            time.sleep(1)

            selectTributacao = Select(driver.find_element_by_name('form1:cmbTributacao'))
            selectTributacao.select_by_value('4')
            time.sleep(1)


            selectRecolhimento = Select(driver.find_element_by_name('form1:cmbRecolhimento'))
            selectRecolhimento.select_by_value('1')
            time.sleep(2)

            #driver.find_element_by_name('form1:j_id433').click() #indo para detalhamento da nota
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "form1:j_id433"))).click()
            print ("Passo 3 - Detalhamento de Nota")
            time.sleep(1)


            descrNota1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:informacoesComplementares")))
            #descrNota1 = driver.find_element_by_id("form1:informacoesComplementares")
            descrNota1.send_keys("SVA - SERVIÇO DE VALOR ADICIONADO")
            print ("Descrição 1 da Nota inserida")

            descrNota2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:descricaoItem")))
            #descrNota2 = driver.find_element_by_id("form1:descricaoItem")
            descrNota2.send_keys("SVA")
            print ("Descrição 2 da Nota inserida")

            #qntItem = driver.find_element_by_id("form1:quantidadeItem")
            qntItem = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:quantidadeItem")))
            qntItem.send_keys("1")
            print ("Quantidade de Itens inserida")

            #valUnitario = driver.find_element_by_id("form1:valorUnitarioItem")
            valUnitario = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "form1:valorUnitarioItem")))
            valUnitario.send_keys(dict["form1:valorUnitarioItem"])
            print ("Valor Unitário Inserido")

            #driver.find_element_by_name('form1:j_id604').click() #adiciona item
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "form1:j_id604"))).click()
            print ("Item adicionado com sucesso")
            time.sleep(2)



            #geraNota = driver.find_elements_by_xpath("//*[contains(text(), ' Visualizar Nota')]")
            #geraNota = driver.find_elements_by_xpath("//*[contains(text(), ' Emitir Nota Fiscal')]")
            #for btn in geraNota:
            #    btn.click()
            #time.sleep(1)


            #driver.find_element_by_name('j_id774:j_id796').click()
            time.sleep(1)
            #driver.get_screenshot_as_file("teste.png")
            print("Cliente {} :".format(i))
            print ("Nota do Cliente : {} gerada com sucesso".format(dict["form1:razaoSocialTomador"]))
            driver.get("https://stm.semfaz.saoluis.ma.gov.br/sistematributario/jsp/login/bemVindo.jsf")

            i = i + 1
            if (i == 500):
                sys.exit()
        #except NoSuchElementException:
        #    print ("form não encontrado para emissão do cliente {}".format(dict["form1:razaoSocialTomador"]))
        #    sys.exit()
        except ElementNotSelectableException:
            print ("não foi possivel selecionar o elemento do cliente {}".format(dict["form1:razaoSocialTomador"]))
            print(ElementNotSelectableException)
            sys.exit()
        except ElementNotVisibleException:
            print ("elemento não visível para o cliente {}".format(dict["form1:razaoSocialTomador"]))
            sys.exit()


