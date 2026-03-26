from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
import os
import unicodedata
from tkinter import Tk, filedialog
import time
import glob
from datetime import datetime
from pathlib import Path
import win32com.client
import configparser
import shutil
import pandas as pd
import re
from selenium.common.exceptions import TimeoutException

APP_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".RoboMediador")
os.makedirs(APP_CONFIG_DIR, exist_ok=True)
config_file_path = os.path.join(APP_CONFIG_DIR, "config.ini")

execucao_registrada = False
caminho_log_hist = None

################################################################

def stop_execution():
    global _stop_flag, _browser_driver
    _stop_flag = True
    if _browser_driver:
        try:
            _browser_driver.quit()
        except Exception as e:
            print(f"Erro ao fechar o navegador: {e}")

################################################################

def load_last_excel_path():
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path)
        if 'Settings' in config and 'LastExcelPath' in config['Settings']:
            return config['Settings']['LastExcelPath']
    return None

################################################################

def save_last_excel_path(path):
    config = configparser.ConfigParser()
   
    if os.path.exists(config_file_path):
        config.read(config_file_path)

    if 'Settings' not in config:
        config['Settings'] = {}

    config['Settings']['LastExcelPath'] = path
    with open(config_file_path, 'w') as configfile:
        config.write(configfile)

################################################################

def log_erro(arquivo_log, mensagem, cnpj=None):
    data_hora_atual = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    with open(arquivo_log, 'a') as log:
        log.write(f"{data_hora_atual}: {mensagem}\n")
        if cnpj:
            log.write(f"CNPJ: {cnpj}\n")
    print(f"Erro registrado em {arquivo_log}: {mensagem}")

################################################################

def load_last_work_path():
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path)
        if 'Settings' in config and 'LastWorkPath' in config['Settings']:
            return config['Settings']['LastWorkPath']
    return None

################################################################

def save_last_work_path(path):
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path)
    if 'Settings' not in config:
        config['Settings'] = {}
    config['Settings']['LastWorkPath'] = path
    with open(config_file_path, 'w') as configfile:
        config.write(configfile)

################################################################

def remove_acentos(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')

################################################################

def obter_caminho_arquivo_excel():
    print("Selecione o arquivo do Excel:")
    root = Tk()
    root.withdraw()  # Esconder a janela principal do Tkinter

    last_path = load_last_excel_path()

    initial_dir = os.path.dirname(last_path) if last_path and os.path.exists(last_path) else os.getcwd()

    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo do Excel",
        initialdir=initial_dir,
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsb")],
    )
    if file_path:
        save_last_excel_path(file_path) 

    return file_path

################################################################

def primeira_letra(palavra):
    return palavra[0].upper()

################################################################

def renomear_arquivo_baixado(codsind, municipio, cidade_sind, numloja, usar_pasta_municipio=False): # Remover dias_de_atraso e sufixo_arquivo
    if _stop_flag: return False
    max_tentativas = 3
    tentativa = 0
    
    #mudança 5 - nova regex para extrair a parte da cidade sem números ou hífens, considerando possíveis variações de formatação
    if cidade_sind: 
        match = re.search(r"-\s*([^0-9-]+?)(?:\s*-|\s+\d|$)", cidade_sind)
        if match:
            cidade_sind_limpa = match.group(1).strip().upper()
        else:
            cidade_sind_limpa = cidade_sind.replace("'", "").strip().upper()
    else:
        cidade_sind_limpa = "aa"

    novo_caminho = None
    caminho_final_pdf = None 
    pasta_municipio_para_retorno = None 
    
    while tentativa < max_tentativas:
        if _stop_flag: return False
        time.sleep(3)  
        
        list_of_files = glob.glob(os.path.join(download_dir, "*.doc"))
        
        if not list_of_files:
            print(f"Nenhum arquivo DOC encontrado (tentativa {tentativa + 1}/{max_tentativas})")
            tentativa += 1
            continue
            
        latest_file = max(list_of_files, key=os.path.getctime)
        print(f"Arquivo encontrado: {latest_file}")
        
        try:
            size1 = os.path.getsize(latest_file)
            time.sleep(1)
            size2 = os.path.getsize(latest_file)
            
            if size1 != size2: 
                print(f"Arquivo ainda sendo baixado (tentativa {tentativa + 1}/{max_tentativas})")
                tentativa += 1
                continue
                
            if "__" in str(numloja):
                numloja_base, sufixo_str = str(numloja).split("__")
                numloja_formatado = str(numloja_base).zfill(3)
                sufixo_formatado = f"__{sufixo_str}"
            else:
                numloja_formatado = str(numloja).zfill(3)
                sufixo_formatado = "" 

            numloja_para_pasta = numloja_formatado 

            codsind_formatado = str(codsind).zfill(4)
            municipio_sem_acentos = remove_acentos(municipio)
            municipio_formatado = municipio_sem_acentos.upper()
            novo_nome_doc = f"{numloja_formatado}_{municipio_formatado}_{cidade_sind_limpa}_SIND{codsind_formatado}{sufixo_formatado}.doc"


            if usar_pasta_municipio:
                pasta_municipio_para_retorno = subpasta_municipio(numloja_para_pasta, municipio)
                novo_caminho = os.path.join(pasta_municipio_para_retorno, novo_nome_doc)
                caminho_final_pdf = novo_caminho.replace(".doc", ".pdf") 
            else:
                novo_caminho = os.path.join(download_dir, novo_nome_doc)
                caminho_final_pdf = novo_caminho.replace(".doc", ".pdf")
            
            try:
                if os.path.exists(novo_caminho):
                    os.remove(novo_caminho)
                
                os.rename(latest_file, novo_caminho)
                print(f"Arquivo renomeado para: {novo_nome_doc}")

                if conversao_doc_para_pdf(novo_caminho):

                    return True, caminho_final_pdf, pasta_municipio_para_retorno 
                
                else:
                    print("Conversão para PDF falhou, mantendo arquivo .doc")
                    return False, None, None 
            
            except PermissionError:
                print("Permissão negada, tentando novamente...")
                tentativa += 1
                continue
            except Exception as e:
                log_erro(arquivo_log, f"Erro ao renomear arquivo: {str(e)}")
                return False, None, None 
                
        except Exception as e:
            log_erro(arquivo_log, f"Erro ao verificar arquivo: {str(e)}")
            tentativa += 1
            continue
    
    log_erro(arquivo_log, f"Não foi possível renomear o arquivo após {max_tentativas} tentativas")
    return False, None, None 

################################################################

def organizar_historicos(diretorio):
    try:
        arquivos = list(Path(diretorio).glob("histórico de buscas_*.txt"))

        if len(arquivos) <= 30:
            return
        
        arquivos.sort(key=lambda f: f.stat().st_mtime, reverse=True)

        for arquivo in arquivos[30:]:
            try:
                os.remove(arquivo)
                print(f"Removido: {arquivo}")
            except Exception as erro:
                print(f"Erro ao remover {arquivo}: {erro}")

    except Exception as erro_geral:
        print(f"Erro ao gerenciar arquivos do histórico: {erro_geral}")

################################################################

def registrar_ultimo_historico():
    global execucao_registrada, caminho_log_hist

    if not execucao_registrada:

        data_hist = datetime.now().strftime('%d-%m-%Y_%H-%M')
        caminho_log_hist = os.path.join(historico_dir, f"histórico de buscas_{data_hist}.txt")

        os.makedirs(os.path.dirname(caminho_log_hist), exist_ok=True)

        ultima_execucao = f"### Execução iniciada em: {datetime.now().strftime('%d/%m/%Y às %H:%M')} ###\n\n"

        with open(caminho_log_hist, "a", encoding="utf-8") as arquivo:
            if os.path.getsize(caminho_log_hist) > 0:
                arquivo.write("\n" + "*" * 60 + "\n")
            arquivo.write(ultima_execucao)

        execucao_registrada = True 

        organizar_historicos(os.path.dirname(caminho_log_hist))

################################################################

def registrar_sem_doc(numloja, cnpj, municipio, codsind):
    try:
        registrar_ultimo_historico()

        data_atual = datetime.now().strftime("%d/%m/%Y às %H:%M")
        linha = f"Loja: {numloja} | Código do sindicato: {codsind} | CNPJ: {cnpj} | Município: {municipio} | Nenhum documento foi encontrado em: {data_atual}\n"

        if os.path.exists(caminho_log_hist):
            with open(caminho_log_hist, "r", encoding="utf-8") as arquivo:
                conteudo = arquivo.readlines()
        else:
            conteudo = []

        conteudo.append(linha)

        with open(caminho_log_hist, "w", encoding="utf-8") as arquivo:
            arquivo.writelines(conteudo)

        print("Loja registrada no .txt")

    except Exception as erro_txt:
        print(f"Erro ao registrar loja no txt: {erro_txt}")

################################################################

def subpasta_municipio(numloja, municipio):
    municipio_sem_acentos = remove_acentos(municipio).upper()
    numloja = str(numloja).zfill(3)
    nome_pasta = f"{numloja}_{municipio_sem_acentos}"
    caminho_pasta = os.path.join(download_dir, nome_pasta)
    os.makedirs(caminho_pasta, exist_ok=True)
    return caminho_pasta

################################################################

def conversao_doc_para_pdf(caminho_doc): # Mudança 6 // talvez refazer
    if _stop_flag: return False
    word = None
    doc = None

    try:
        word = win32com.client.DispatchEx("Word.Application")
        
        word.Visible = False
        word.DisplayAlerts = 0 
        word.ScreenUpdating = False 

        caminho_doc_absoluto = os.path.abspath(caminho_doc)
        print(f"INFO: Convertendo em segundo plano (processo isolado): {caminho_doc_absoluto}")
        doc = word.Documents.Open(FileName=caminho_doc_absoluto, ReadOnly=True, Visible=False)
        
        caminho_pdf = os.path.splitext(caminho_doc_absoluto)[0] + ".pdf"
        doc.SaveAs(caminho_pdf, FileFormat=17)
        
        doc.Close(False)
        doc = None

        word.Quit()
        word = None

        if os.path.exists(caminho_pdf) and os.path.exists(caminho_doc_absoluto):
            os.remove(caminho_doc_absoluto)
         
        return True 

    except Exception as e:
        if 'arquivo_log' in globals():
            log_erro(arquivo_log, f"Erro ao converter DOC para PDF: {str(e)}")
        
        if doc:
            try: doc.Close(False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        return False
################################################################   # Função nova adicionada para aguardar elementos de forma mais robusta,
                                                                   # evitando erros comuns de timing e carregamento do site.
                                                                   # Substitui os trechos de código repetitivos de espera e clique/preenchimento 
                                                                   # no def movimentacao_site, deixando o código mais limpo e fácil de manter.

def aguardar_e_clicar(browser, xpath, tempo=30): # Mudança 7 
    try:
        elemento = WebDriverWait(browser, tempo).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        elemento.click()
        return elemento
    except TimeoutException:
        print(f"Erro: Tempo esgotado esperando o elemento: {xpath}")
        return None

def aguardar_e_preencher(browser, xpath, texto, tempo=15):
    try:
        elemento = WebDriverWait(browser, tempo).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        elemento.clear() # Garante que o campo esteja vazio
        elemento.send_keys(texto)
        return elemento
    except TimeoutException:
        print(f"Erro: Tempo esgotado para preencher: {xpath}")
        return None

################################################################

def movimentacao_site(cnpj, browser, uf, municipio, primeira_letra_municipio, codsind, numloja, cidade_sind, dias_de_atraso):
    global _stop_flag
    if _stop_flag: return
    try:
        time.sleep(2)
        browser.refresh()
        time.sleep(3)
        if _stop_flag: return

        print("Inserindo informações no site e iniciando as buscas")
        
        #cnpj checkbox
        aguardar_e_clicar(browser, '//*[@id="chkNRCNPJ"]')
        if _stop_flag: return

        #cnpj input
        aguardar_e_preencher(browser, '//*[@id="txtNRCNPJ"]', cnpj)
        if _stop_flag: return

        print("Selecionando opção da categoria")
        #categoria checkbox
        aguardar_e_clicar(browser, '//*[@id="chkCategoria"]')
        if _stop_flag: return

        print("Inserindo categoria 'Comercio'")
        #categoria input
        aguardar_e_preencher(browser, '//*[@id="txtDSCategoria"]', 'Comercio')
        if _stop_flag: return

        print("Selecionando o Tipo de Instrumento")
        #tipo instrumento
        aguardar_e_clicar(browser, '//*[@id="cboTPRequerimento"]/option[5]')
        if _stop_flag: return

        print("Selecionando a Vigência")
        #vigencia
        aguardar_e_clicar(browser, '//*[@id="cboSTVigencia"]/option[3]')
        if _stop_flag: return

        print("Selecionando a primeira UF")
        #uf registro
        aguardar_e_clicar(browser, '//*[@id="cboUFRegistro"]')
        #uf opção registro
        aguardar_e_clicar(browser, f"//*[@id='cboUFRegistro']/option[@value='{uf}']")
        if _stop_flag: return

        print("Selecionando a segunda UF")
        #abrangencia territorial
        aguardar_e_clicar(browser, '//*[@id="cboUFAbrangenciaTerritorial"]')
        #abrangencia opção
        aguardar_e_clicar(browser, f"//*[@id='cboUFAbrangenciaTerritorial']/option[@value='{uf}']")
        if _stop_flag: return
        
        #selecionar município
        aguardar_e_clicar(browser, '//*[@id="btnSelecionarMunicipio"]')
        if _stop_flag: return

        time.sleep(1)
        print("Filtrando a primeira letra do município")
        if primeira_letra_municipio.upper() != "A":    
            letra_xpath = f'//div[@id="divMunicipiosAbrangenciaTerritorial"]//td[contains(text(), "{primeira_letra_municipio}")]'
        
            aguardar_e_clicar(browser, letra_xpath)
        else:
            print("Seleção já começa na letra A, pulando o clique....")
        if _stop_flag: return

        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.ID, "divMunicipiosSelecionados"))
        )
        if _stop_flag: return

        print("Selecionando o município")
        time.sleep(1.3)
        municipio_normalizado = remove_acentos(municipio.lower())
        municipios_xpath = '//*[@id="divMunicipiosSelecionados"]//label'
        municipios_elementos = WebDriverWait(browser, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, municipios_xpath)))
        if _stop_flag: return

        time.sleep(1.5)
        try:
            municipio_encontrado = False
            for elemento in municipios_elementos:
                if _stop_flag: return
                if remove_acentos(elemento.text.strip().lower()) == municipio_normalizado:
                    elemento.click()
                    municipio_encontrado = True
                    break
            
            if not municipio_encontrado:
                raise ValueError(f"Município '{municipio}' não encontrado na lista do site.")

            aguardar_e_clicar(browser, '/html/body/div[3]/div[3]/div/button/span')
            if _stop_flag: return
        except Exception as e:
            print("Ocorreu um erro ao selecionar o municipio")
            log_erro(arquivo_log, f"Erro no município CNPJ: {cnpj}: {str(e)}")
            return
    
        time.sleep(1)
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        print("Clicando no botão de pesquisa")
        aguardar_e_clicar(browser, '//*[@id="btnPesquisar"]')
        if _stop_flag: return

        ultimo_caminho_pdf_processado = None
        pasta_municipio_final = None 
        
        try:
            print("aguardandoalerta de nenhum registro detectado") #1
            alerta = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[2]/table/tbody/tr/td[2]')))
            
            print("Alerta de nenhum registro detectado")
            aguardar_e_clicar(browser, '/html/body/div[5]/div[3]/div/button/span')
            
            registrar_sem_doc(numloja, cnpj, municipio, codsind)

            aguardar_e_clicar(browser, '//*[@id="btnLimpar"]')
            aguardar_e_clicar(browser, '/html/body/div[5]/div[3]/div/button[1]/span')

            print(f"Registrando loja sem documento: {numloja}")
            registrar_sem_doc(numloja, cnpj, municipio, codsind)
            if _stop_flag: return
            return
        
        except Exception: 
            print("Nenhum alerta detectado, seguindo normalmente para o download")
            browser.execute_script("window.scrollTo(0, 0);")
            if _stop_flag: return
            
            original_window = browser.current_window_handle
        
            print("Procurando todos os botões de download...")
            botoes_download = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')

            print("Procurando todos os botões de download...")
            botoes_download = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
            if _stop_flag: return

            if not botoes_download:
                print("Nenhum botão de download encontrado após a pesquisa. Isso é inesperado.")
                registrar_sem_doc(numloja, cnpj, municipio, codsind)
            else:
                print(f"{len(botoes_download)} botões de download encontrados")

                usar_pasta_municipio_na_chamada = len(botoes_download) > 1

                
                for i in range(len(botoes_download)):
                    if _stop_flag: return

                    try:   
                        botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                        if i >= len(botoes_atuais): break

                        botao_foco = botoes_atuais[i]
                        pular_download = False

                        print(f"Processando item {i+1} de {len(botoes_download)}...")

                        container_pai = botao_foco.find_element(By.XPATH, './../..')
                        botao_visualizar = container_pai.find_element(By.XPATH, ".//a[contains(normalize-space(.), 'Visualizar Instrumento Coletivo')]")

                        #
                        janela_original = browser.current_window_handle
                        janelas_antes = set(browser.window_handles)

                        browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_visualizar)
                        time.sleep(1)
                        browser.execute_script("arguments[0].click();", botao_visualizar)

                        WebDriverWait(browser, 15).until(EC.number_of_windows_to_be(len(janelas_antes) + 1))
                        nova_aba = (set(browser.window_handles) - janelas_antes).pop()
                        browser.switch_to.window(nova_aba)
                        
                        time.sleep(5)

                        texto_pagina = browser.find_element(By.TAG_NAME, "body").text.lower()

                            # arrumar
                        palavras_para_buscar = ["rodoviária", 
                                                          "rodov", "peix", "peixarias",
                                                          "rodoviarias", "ATACADISTA", "FARMACEUTICOS", 
                                                          "OPTICO", "ATAC DE AL E BEB EM GERAL",
                                                          "HIPERMERCADOS", "COMERCIO DE BENS E DE SERVICOS",
                                                          "ORGANIZACAO DAS COOPERATIVAS",
                                                          "COMERCIO DE VEICULOS",
                                                          "ESTABELECIMENTOS DE PRESTACAO DE SERVICOS FUNERARIOS",
                                                          "SINDICATO DOS LOJISTAS",
                                                          "MAT.CONSTR. LOUCAS, TINTAS, FERRAG.",
                                                          "ATAC DE PRODUTOS QUIM", "OCERGS-SINDICATO E ORGANIZACAO",
                                                          "CONCESSIONARIOS", "DISTRIBUIDORES DE VEICULOS", "11.564.609/0001-00", "90.818.667/0001-99", "90.818.667/0001-99",
                                                          "90.818.667/0001-99", "ORGANIZACAO DAS COOPERATIVAS DO ESTADO"]    # mudança 3
                             
                        # 11.564.609/0001-00 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS DO VALE DO SINOS - SINDIGENEROS/VALE
                        # 90.818.667/0001-99 - SINDICATO INTERMUNICIPAL DO COM. VAREJ. DE GEN. ALIM. E DO C
                        # 90.818.667/0001-99 - SINDICATO INTERMUNICIPAL DO COM. VAREJ. DE GEN. ALIM. E DO C
                        # 90.818.667/0001-99 - SINDICATO INTERMUNICIPAL DO COM. VAREJ. DE GEN. ALIM. E DO C

                        print("Buscando por palavras-chave na aba de visualização...")
                            
                        time.sleep(8) # mudança 4

                        texto_da_pagina = browser.find_element(By.TAG_NAME, "body").text
                        texto_da_pagina_lower = texto_da_pagina.lower()

                        for palavra in palavras_para_buscar:
                            if palavra.lower() in texto_da_pagina_lower:

                                print("          ")
                                print(f"  AVISO: Palavra-chave '{palavra}' encontrada para a Loja {numloja}.")
                                print("          ")
                                pular_download = True
                                break

                        if not pular_download:

                            botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                            botao_clique_real = botoes_atuais[i]
                            
                            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_clique_real)
                            time.sleep(1)
                            browser.execute_script("arguments[0].click();", botao_clique_real)
                            print(f"Download disparado para item {i+1}")
                            
                            try:
                                sucesso, pdf_path, folder_path = renomear_arquivo_baixado(
                                    codsind,
                                    municipio,
                                    cidade_sind,
                                    f"{numloja}__{i+1}" if usar_pasta_municipio_na_chamada else numloja,
                                    usar_pasta_municipio=usar_pasta_municipio_na_chamada
                                )

                                if sucesso and dias_de_atraso < 0:
                                    print(f"Documento da Loja {numloja} está atrasado. Movendo para pasta específica...")

                                    if usar_pasta_municipio_na_chamada and folder_path:
                                        destino_atrasado = os.path.join(atrasados_dir, os.path.basename(folder_path))
                                        if os.path.exists(destino_atrasado):
                                            shutil.rmtree(destino_atrasado)
                                        shutil.move(folder_path, atrasados_dir)
                                        print(f"Pasta movida para: {atrasados_dir}")

                                    elif pdf_path:
                                        destino_atrasado = os.path.join(atrasados_dir, os.path.basename(pdf_path))

                                        if os.path.exists(destino_atrasado):
                                            os.remove(destino_atrasado)
                                        shutil.move(pdf_path, atrasados_dir)
                                        print(f"Arquivo movido para: {atrasados_dir}")
                                    
                            except Exception as e_proc:
                                print(f"Erro ao processar arquivo baixado: {e_proc}")

                            #-----------------------------------------------------
  
                            print(f"Clique de download disparado para o item {i+1}")

                        browser.close() # Fecha a aba de visualização (que estava em foco)
                        browser.switch_to.window(original_window) # Volta para a aba da lista
                        time.sleep(1)
                      
                        if pular_download:
                                continue

                        janelas_antes_download = set(browser.window_handles)
                        botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                        botao_clique_real = botoes_atuais[i]

                        browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_clique_real)
                        time.sleep(1.5)
                        browser.execute_script("arguments[0].click();", botao_clique_real)

                        print(f"Download disparado para item {i+1}")
                       

                        try:
                            WebDriverWait(browser, 5).until(lambda drv: len(drv.window_handles) > len(janelas_antes_download))
                            novas_janelas = set(browser.window_handles) - janelas_antes_download
                            for aba in novas_janelas:
                                browser.switch_to.window(aba)
                                browser.close()
                            browser.switch_to.window(janela_original)
                        except:
                            pass

                        #
                        sucesso, pdf_path, folder_path = renomear_arquivo_baixado(
                                            codsind, 
                                            municipio, 
                                            cidade_sind, 
                                            f"{numloja}__{i+1}" if usar_pasta_municipio_na_chamada else numloja, 
                                            usar_pasta_municipio=usar_pasta_municipio_na_chamada
                                        )
                        if sucesso:
                            ultimo_caminho_pdf_processado = pdf_path
                            pasta_municipio_final = folder_path
                        else:
                            print("Não foi possível renomear ou converter um dos arquivos.")

                    except Exception as e:
                        log_erro(arquivo_log, f"Erro ao tentar baixar e processar arquivo {i+1} de {len(botoes_download)} para Loja {numloja}: {str(e)}")

            time.sleep(2)
            if _stop_flag: return

            if isinstance(dias_de_atraso, (int, float)) and dias_de_atraso < 0:
                print(f"Loja {numloja} está atrasada. Movendo todos os documentos para a pasta ATRASADOS.")
                
                if usar_pasta_municipio_na_chamada and pasta_municipio_final:
                    caminho_destino_atrasados_pasta = os.path.join(atrasados_dir, os.path.basename(pasta_municipio_final))
                    
                    if os.path.exists(caminho_destino_atrasados_pasta):
                        try:
                            shutil.rmtree(caminho_destino_atrasados_pasta) 
                            print(f"Pasta '{os.path.basename(pasta_municipio_final)}' existente em ATRASADOS removida para atualização.")
                        except Exception as e:
                            log_erro(arquivo_log, f"Erro ao remover pasta existente em ATRASADOS '{os.path.basename(pasta_municipio_final)}': {str(e)}")
                            return 
                                    
                    shutil.move(pasta_municipio_final, atrasados_dir) 
                    print(f"Pasta '{os.path.basename(pasta_municipio_final)}' movida para ATRASADOS.")
                elif ultimo_caminho_pdf_processado: 
                    caminho_destino_atrasados_pdf = os.path.join(atrasados_dir, os.path.basename(ultimo_caminho_pdf_processado))
                    
                    if os.path.exists(caminho_destino_atrasados_pdf):
                        try:
                            os.remove(caminho_destino_atrasados_pdf)
                            print(f"Arquivo '{os.path.basename(ultimo_caminho_pdf_processado)}' existente em ATRASADOS removido para atualização.")
                        except Exception as e:
                            log_erro(arquivo_log, f"Erro ao remover arquivo PDF existente em ATRASADOS '{os.path.basename(ultimo_caminho_pdf_processado)}': {str(e)}")
                            return
                    
                    shutil.move(ultimo_caminho_pdf_processado, atrasados_dir)
                    print(f"Arquivo '{os.path.basename(ultimo_caminho_pdf_processado)}' movido para ATRASADOS.")
                else:
                    print(f"Loja {numloja} está atrasada, mas nenhum documento foi processado para mover.")

            browser.get('http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo')
            if _stop_flag: return
        except Exception as e:
            log_erro(
                arquivo_log, f"Erro ao baixar o documento para CNPJ: {cnpj}: {str(e)}")
            print(f"Erro ao baixar o documento: {str(e)}")

    except Exception as e:
        log_erro(
            arquivo_log, f"Erro na função movimentacao_site: {str(e)}", cnpj)
        print(f"Erro na função movimentacao_site: {str(e)}")
        
#############################################################################

def main(caminho_excel=None, base_path_param=None, loja_inicial=None, update_callback=None): #  parte principal do código
    global _stop_flag, _browser_driver
    global mediador_dir, historico_dir, download_dir, atrasados_dir
    global pasta_log, arquivo_log, diretorio_log, caminho_log_hist

    _stop_flag = False
    _browser_driver = None

    if base_path_param:
        print(f"Diretório de trabalho definido pela interface: {base_path_param}")
        mediador_dir = base_path_param
        historico_dir = os.path.join(mediador_dir, "Histórico")
        download_dir = os.path.join(mediador_dir, "Downloads")
        atrasados_dir = os.path.join(download_dir, "ATRASADOS")
        pasta_log = os.path.join(mediador_dir, "Logs")
        arquivo_log = os.path.join(pasta_log, "Log Erros.txt")
        diretorio_log = os.path.dirname(arquivo_log)
        caminho_log_hist = None
        
        os.makedirs(diretorio_log, exist_ok=True)
        os.makedirs(download_dir, exist_ok=True)
        os.makedirs(atrasados_dir, exist_ok=True)
        os.makedirs(historico_dir, exist_ok=True)

    try:
        if not caminho_excel:
            caminho_excel = obter_caminho_arquivo_excel()
            if not caminho_excel:
                print('Nenhum arquivo selecionado, encerrando programa.')
                return
        df = pd.read_excel(caminho_excel)
        print("Arquivo Excel carregado")
    except Exception as e:
        log_fallback = "log_erro_inicializacao.txt"
        log_erro(arquivo_log if 'arquivo_log' in locals() else log_fallback, f'Erro ao carregar o arquivo Excel: {e}')
        return

    loja_final_para_iniciar = None
    iniciar_processamento_agora = not bool(loja_inicial)


    todas_as_lojas = []
    loja_inicial_num = None

    if not iniciar_processamento_agora:
        try:
            loja_inicial_num = int(loja_inicial)
            print(f"Procurando por loja inicial a partir de: {loja_inicial_num}")
           
            for row in df.itertuples(index=False):   # mudança 2 junto a troca de openpyxl para pandas
                numloja = row[0]          
                try:
                    if row[0] is not None:
                        todas_as_lojas.append(int(row[0]))
                except (ValueError, TypeError):
                    continue
            
            todas_as_lojas = sorted(list(set(todas_as_lojas)))
            
            lojas_posteriores = [lj for lj in todas_as_lojas if lj >= loja_inicial_num]
            
            if lojas_posteriores:
                loja_final_para_iniciar = min(lojas_posteriores)
                if loja_final_para_iniciar != loja_inicial_num:
                    print(f"AVISO: Loja '{loja_inicial}' não encontrada. Iniciando na loja mais próxima: '{loja_final_para_iniciar}'")
            else:
                print(f"AVISO: Nenhuma loja encontrada no Excel a partir do número '{loja_inicial}'. O processamento não será iniciado.")
                if update_callback:
                    update_callback(f"Loja '{loja_inicial}' não encontrada na planilha.")
                return
        except (ValueError, TypeError):
            print(f"Valor '{loja_inicial}' inválido para loja inicial. Iniciando do começo.")
            iniciar_processamento_agora = True
            
    try:
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("disable-infobars")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--headless=new")
        
        final_download_path = os.path.abspath(download_dir)
        print(f"INFO: Configurando diretório de download do Chrome para: {final_download_path}")
        prefs = {
            "download.prompt_for_download": False, "download.directory_upgrade": True,
            "safeBrowse.enabled": False, "profile.default_content_settings.popups": 0,
            "download.default_directory": final_download_path,
            "safeBrowse.disable_download_protection": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        _browser_driver = webdriver.Chrome(options=chrome_options)
        browser = _browser_driver
    except Exception as e:
        log_erro(arquivo_log, f"Ocorreu um erro ao abrir o Chrome: {str(e)}")
        return

    if browser is not None:
        print("Navegador Iniciado")
        try:
            browser.get('http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo')
            print("Página carregada")

            for row in df.itertuples(index=False):
                numloja = row[0]
                try:
                    if row[0] is not None:
                        todas_as_lojas.append(int(row[0]))
                except (ValueError, TypeError):
                    continue
                if _stop_flag:
                    print("Execução interrompida pela interface.")
                    break
                
                numloja_atual = str(row[0])
                if not iniciar_processamento_agora:
                    if numloja_atual == str(loja_final_para_iniciar):
                        print(f"--- Loja alvo '{loja_final_para_iniciar}' encontrada. Iniciando processamento. ---")
                        iniciar_processamento_agora = True
                    else:
                        print(f"Pulando loja: {numloja_atual}...")
                        continue
                
                if numloja_atual == 'None': continue
                
                # indices sindicatos(edu)
                #cnpj = str(row[2])         
                #uf = str(row[4])
                #municipio = str(row[21])
                #cidade_sind = str(row[20])
                #codsind = str(row[1])
                #numloja = str(row[0])
                
                # indices acompanhamentos
                cnpj = str(row[2])
                uf = str(row[4])
                municipio = str(row[5])
                cidade_sind = str(row[9])
                codsind = str(row[1])
                numloja = str(row[0]) 

                dias_de_atraso = 0
                if len(row) > 3 and row[3] is not None:
                    try:
                        dias_de_atraso = float(row[3])     
                    except (ValueError, TypeError):
                        dias_de_atraso = 0
                primeira_letra_municipio = primeira_letra(municipio) # mudança 1

                if update_callback:
                    update_callback(f"{numloja} | CNPJ: {cnpj}")
                print("\n" + f"Processando: Loja {numloja} | CNPJ {cnpj} | Código Sindicato {codsind} | Município {municipio}" + "\n")

                if not cnpj or cnpj == 'None':
                    print(f"CNPJ ausente ou inválido, pulando: CNPJ lido: {cnpj}")
                    continue
                # chama o selenium para o site
                movimentacao_site(cnpj, browser, uf, municipio, primeira_letra_municipio, codsind, numloja, cidade_sind, dias_de_atraso)

            if not _stop_flag:
                print("Processamento concluído.")

        except Exception as e:
            log_erro(arquivo_log, f"Ocorreu um erro na execução principal: {str(e)}", cnpj=locals().get('cnpj'))
        finally:
            if browser:
                browser.quit()