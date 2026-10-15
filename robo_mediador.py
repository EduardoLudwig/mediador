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
from selenium.common.exceptions import InvalidSessionIdException
from selenium.common.exceptions import StaleElementReferenceException
import subprocess
from enviar_email import enviar_relatorio

APP_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".RoboMediador")
os.makedirs(APP_CONFIG_DIR, exist_ok=True)
config_file_path = os.path.join(APP_CONFIG_DIR, "config.ini")
# Pasta de logs que a limpeza pós-email NÃO apaga (fica fora da pasta de trabalho, junto do config.ini dentro da pasta .RoboMediador dentro do usuario no C)
PASTA_LOG_PERMANENTE = os.path.join(APP_CONFIG_DIR, "Logs Permanentes")
os.makedirs(PASTA_LOG_PERMANENTE, exist_ok=True)

################################################################

def garantir_config_inicial():
    # Cria o config.ini automaticamente na primeira execução (sem precisar selecionar nada na interface),
    # já com as seções [Settings] e [Email] definidas. Se o arquivo já existir, NÃO sobrescreve nada:
    # preserva os valores preenchidos manualmente e só adiciona o que estiver faltando (ex.: a seção [Email]).
    config = configparser.ConfigParser()
    alterado = False

    if os.path.exists(config_file_path):
        config.read(config_file_path, encoding='utf-8')

    if 'Settings' not in config:
        config['Settings'] = {}
        alterado = True

    if 'Email' not in config:
        config['Email'] = {}
        alterado = True

    email_cfg = config['Email']
    for chave in ('EmailRemetente', 'SenhaApp', 'EmailDestinatario'):
        if chave not in email_cfg:
            email_cfg[chave] = ''
            alterado = True

    if alterado:
        with open(config_file_path, 'w', encoding='utf-8') as configfile:
            config.write(configfile)

################################################################

garantir_config_inicial()

execucao_registrada = False
caminho_log_hist = None
_falhas_consecutivas = {}

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
        config.read(config_file_path, encoding='utf-8')
        if 'Settings' in config and 'LastExcelPath' in config['Settings']:
            return config['Settings']['LastExcelPath']
    return None

################################################################

def save_last_excel_path(path):
    config = configparser.ConfigParser()
   
    if os.path.exists(config_file_path):
        config.read(config_file_path, encoding='utf-8')

    if 'Settings' not in config:
        config['Settings'] = {}

    config['Settings']['LastExcelPath'] = path
    with open(config_file_path, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

################################################################

def log_erro(arquivo_log, mensagem, cnpj=None, numloja=None):
    data_log = datetime.now().strftime("%d-%m-%Y_%H-%M")
    layout_mensagem = (
        f"Erro na data: {data_log}\n"
        f"LOJA: {numloja if numloja else 'N/A'} - CNPJ: {cnpj if cnpj else 'N/A'}\n\n"
        f"{mensagem}\n"
        f"\n-------------------------------------------------\n\n"
    )
    # Grava o erro em dois lugares: no Logs da pasta de trabalho (vai no email e é limpo depois) e no Logs Permanentes em ~/.RoboMediador, com o mesmo nome de arquivo (não é apagado pela limpeza pós-email, guarda os 30 mais recentes, serve para rastreabilidade)
    arquivo_log_permanente = os.path.join(PASTA_LOG_PERMANENTE, os.path.basename(arquivo_log))
    for caminho in (arquivo_log, arquivo_log_permanente):
        with open(caminho, 'a', encoding='utf-8') as log:
            log.write(layout_mensagem)
    print(layout_mensagem)

################################################################

def organizar_Log_erro(diretorio_log):
    try:
        arquivos = list(Path(diretorio_log).glob("Log Erros_*.txt"))

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
        print(f"Erro ao gerenciar arquivos de logs: {erro_geral}")

################################################################

def load_last_work_path():
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path, encoding='utf-8')
        if 'Settings' in config and 'LastWorkPath' in config['Settings']:
            return config['Settings']['LastWorkPath']
    return None

################################################################

def save_last_work_path(path):
    config = configparser.ConfigParser()
    if os.path.exists(config_file_path):
        config.read(config_file_path, encoding='utf-8')
    if 'Settings' not in config:
        config['Settings'] = {}
    config['Settings']['LastWorkPath'] = path
    with open(config_file_path, 'w', encoding='utf-8') as configfile:
        config.write(configfile)

################################################################

def remove_acentos(texto):
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')

################################################################

def obter_caminho_arquivo_excel():
    print("Selecione o arquivo do Excel:")
    root = Tk()
    root.withdraw()

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

def renomear_arquivo_baixado(codsind, municipio, cidade_sind, cnpj, numloja, arquivos_antes, usar_pasta_municipio=False, is_debug=False):
    if _stop_flag: return False, None, None
    # Às vezes o site demora pra entregar o arquivo, por isso 10 tentativas (~40s de espera)
    max_tentativas = 10
    tentativa = 0
    
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
        if _stop_flag: return False, None, None
        time.sleep(3)

        # Só considera arquivos que NÃO estavam na pasta antes do clique. Sem isso, se o download de uma loja atrasasse, o arquivo dela chegava depois e era renomeado com o nome da loja seguinte
        list_of_files = [
            arquivo for arquivo in glob.glob(os.path.join(download_dir, "*.doc"))
            if arquivo not in arquivos_antes
        ]
        
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

            diretorio_destino = filtrados_debug_dir if is_debug else download_dir

            if usar_pasta_municipio:
                pasta_municipio_para_retorno = subpasta_municipio(numloja_para_pasta, municipio, diretorio_base=diretorio_destino)
                novo_caminho = os.path.join(pasta_municipio_para_retorno, novo_nome_doc)
                caminho_final_pdf = novo_caminho.replace(".doc", ".pdf") 
            else:
                novo_caminho = os.path.join(diretorio_destino, novo_nome_doc)
                caminho_final_pdf = novo_caminho.replace(".doc", ".pdf")
            
            try:
                if os.path.exists(novo_caminho):
                    os.remove(novo_caminho)
                
                os.rename(latest_file, novo_caminho)
                print(f"Arquivo renomeado para: {novo_nome_doc}")

                if conversao_doc_para_pdf(novo_caminho, cnpj, numloja):

                    return True, caminho_final_pdf, pasta_municipio_para_retorno 
                
                else:
                    print("Conversão para PDF falhou, mantendo arquivo .doc")
                    return False, None, None 
            
            except PermissionError:
                print("Permissão negada, tentando novamente...")
                tentativa += 1
                continue
            except Exception as e:
                log_erro(arquivo_log, f"Erro ao renomear arquivo: {str(e)}", cnpj, numloja)
                return False, None, None 
                 
        except Exception as e:
            log_erro(arquivo_log, f"Erro ao verificar arquivo: {str(e)}", cnpj, numloja)
            tentativa += 1
            continue
    
    log_erro(arquivo_log, f"Não foi possível renomear o arquivo após {max_tentativas} tentativas (Arquivo pode não ter sido baixado, estar em uso ou a permissão pode ter sido negada.)", cnpj, numloja)
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

def registrar_loja_pulada(numloja, cnpj, municipio, codsind, motivo):
    try:
        registrar_ultimo_historico()

        data_atual = datetime.now().strftime("%d/%m/%Y às %H:%M")
        linha = f"Loja: {numloja} | Código do sindicato: {codsind} | CNPJ: {cnpj} | Município: {municipio} | {motivo} em: {data_atual}\n"

        if os.path.exists(caminho_log_hist):
            with open(caminho_log_hist, "r", encoding="utf-8") as arquivo:
                conteudo = arquivo.readlines()
        else:
            conteudo = []

        conteudo.append(linha)

        with open(caminho_log_hist, "w", encoding="utf-8") as arquivo:
            arquivo.writelines(conteudo)

        print("Loja pulada registrada no .txt")

    except Exception as erro_txt:
        print(f"Erro ao registrar loja pulada no txt: {erro_txt}")

################################################################

def subpasta_municipio(numloja, municipio, diretorio_base=None):
    if diretorio_base is None:
        diretorio_base = download_dir
        
    municipio_sem_acentos = remove_acentos(municipio).upper()
    numloja = str(numloja).zfill(3)
    nome_pasta = f"{numloja}_{municipio_sem_acentos}"
    caminho_pasta = os.path.join(diretorio_base, nome_pasta)
    os.makedirs(caminho_pasta, exist_ok=True)
    return caminho_pasta

################################################################

def verifica_office_instalado():

    program_files = os.environ.get("PROGRAMFILES", "")
    program_files_x86 = os.environ.get("PROGRAMFILES(X86)", "")

    caminhos = [
        os.path.join(program_files, "Microsoft Office", "Office16", "WINWORD.EXE"),
        os.path.join(program_files, "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
        os.path.join(program_files_x86, "Microsoft Office", "Office16", "WINWORD.EXE"),
        os.path.join(program_files_x86, "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
    ]

    for caminho in caminhos:
        if os.path.isfile(caminho):
            return True
    return False
################################################################

def encontra_libreoffice():
    caminhos_comuns = [
        os.path.join(os.environ.get("PROGRAMFILES", ""), "LibreOffice", "program", "soffice.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "LibreOffice", "program", "soffice.exe"),
    ]
    
    for caminho in caminhos_comuns:
        if os.path.isfile(caminho):
            return caminho
    
    try:
        resultado = subprocess.run(
            ["where", "soffice.exe"],
            capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
        )
        if resultado.returncode == 0:
            caminho = resultado.stdout.strip().split("\n")[0]
            if os.path.isfile(caminho):
                return caminho
    except Exception:
        pass
    
    return None

################################################################

def conversao_doc_para_pdf(caminho_doc, cnpj, numloja):
    if _stop_flag: return False
    caminho_doc_absoluto = os.path.abspath(caminho_doc)
    caminho_pdf = os.path.splitext(caminho_doc_absoluto)[0] + ".pdf"
    pasta_destino = os.path.dirname(caminho_doc_absoluto)
    
    # Verifica se Office está instalado antes de tentar Word COM
    if verifica_office_instalado():
        print(f"INFO: Microsoft Office detectado. Convertendo via Word...")
        word = None
        doc = None

        try:
            word = win32com.client.DispatchEx("Word.Application")
            
            word.Visible = False
            word.DisplayAlerts = 0  
            word.ScreenUpdating = False 

            print(f"INFO: Convertendo em segundo plano (processo isolado): {caminho_doc_absoluto}")
            doc = word.Documents.Open(FileName=caminho_doc_absoluto, ReadOnly=True, Visible=False)
            
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
                log_erro(arquivo_log, f"Erro ao converter DOC para PDF via Word: {str(e)}", cnpj, numloja)
            
            if doc:
                try: doc.Close(False)
                except: pass
            if word:
                try: word.Quit()
                except: pass
            # Se Word falhou, continua para LibreOffice abaixo
    
    # Fallback tenta converter via LibreOffice
    else:
        print(f"INFO: Microsoft Office NÃO detectado.")
    
    print(f"INFO: Tentando conversão via LibreOffice...")
    
    caminho_soffice = encontra_libreoffice()
    if caminho_soffice:
        try:
            resultado = subprocess.run(
                [caminho_soffice, "--headless", "--convert-to", "pdf", "--outdir", pasta_destino, caminho_doc_absoluto],
                capture_output=True, text=True, timeout=120,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if resultado.returncode == 0 and os.path.exists(caminho_pdf):
                print(f"INFO: Conversão via LibreOffice concluída com sucesso.")
                if os.path.exists(caminho_doc_absoluto):
                    os.remove(caminho_doc_absoluto)
                return True
            else:
                print(f"ERRO: LibreOffice falhou: {resultado.stderr}")
        except subprocess.TimeoutExpired:
            print(f"ERRO: LibreOffice excedeu tempo limite de conversão.")
        except Exception as e_liber:
            print(f"ERRO: Falha na conversão via LibreOffice: {str(e_liber)}")
    else:
        print(f"ERRO: LibreOffice não encontrado. Instale o LibreOffice para converter DOC para PDF.")
    
    return False
################################################################  

def aguardar_e_clicar(browser, xpath, tempo=30):
    try:
        elemento = WebDriverWait(browser, tempo).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        try:
            elemento.click()
        except Exception:
            # Clique via JS, ignora overlay temporário
            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
            browser.execute_script("arguments[0].click();", elemento)
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
            max_tentativas_stale = 3
            for tentativa_stale in range(max_tentativas_stale):
                try:
                    municipios_elementos = browser.find_elements(By.XPATH, municipios_xpath)
                    for elemento in municipios_elementos:
                        if _stop_flag: return
                        if remove_acentos(elemento.text.strip().lower()) == municipio_normalizado:
                            try:
                                WebDriverWait(browser, 5).until(
                                    EC.invisibility_of_element_located((By.CSS_SELECTOR, ".ui-widget-overlay"))
                                )
                            except TimeoutException:
                                pass
                            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                            browser.execute_script("arguments[0].click();", elemento)
                            municipio_encontrado = True
                            break
                    break  
                except StaleElementReferenceException:
                    print(f"Elemento obsoleto, tentando novamente ({tentativa_stale+1}/{max_tentativas_stale})...")
                    time.sleep(1)
                    continue
            
            if not municipio_encontrado:
                print(f"Município '{municipio}' não encontrado na lista do site.")
                registrar_sem_doc(numloja, cnpj, municipio, codsind)
                return

            aguardar_e_clicar(browser, '/html/body/div[3]/div[3]/div/button/span')
            if _stop_flag: return
        except Exception as e:
            print("Ocorreu um erro ao selecionar o municipio")
            log_erro(arquivo_log, f"Erro no município CNPJ: {cnpj} LOJA: {numloja}: {str(e)}")
            return
    
        time.sleep(1)
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        print("Clicando no botão de pesquisa")
        aguardar_e_clicar(browser, '//*[@id="btnPesquisar"]')
        if _stop_flag: return
        
        try:
            print("aguardando alerta de nenhum registro detectado")
            
            WebDriverWait(browser, 5).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[2]/table/tbody/tr/td[2]')))
            
            print("Alerta de nenhum registro detectado")
            aguardar_e_clicar(browser, '/html/body/div[5]/div[3]/div/button/span')
            
            registrar_sem_doc(numloja, cnpj, municipio, codsind)

            aguardar_e_clicar(browser, '//*[@id="btnLimpar"]')
            aguardar_e_clicar(browser, '/html/body/div[5]/div[3]/div/button[1]/span')

            if _stop_flag: return
            return
        
        except Exception: 
            print("Nenhum alerta detectado, seguindo normalmente para o download")
            browser.execute_script("window.scrollTo(0, 0);")
            if _stop_flag: return
            
            original_window = browser.current_window_handle
        
            print("Procurando todos os botões de download na tela...")
            botoes_download_iniciais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
            if _stop_flag: return

            if not botoes_download_iniciais:
                print("Nenhum botão de download encontrado após a pesquisa.")
                registrar_sem_doc(numloja, cnpj, municipio, codsind)
            else:
                total_botoes = len(botoes_download_iniciais)
                print(f"{total_botoes} botões de download encontrados. Iniciando varredura...")
                
                indices_aprovados = [] # Lista que guardará os botões que passaram no filtro
                documentos_reprovados = [] # Guardará os reprovados no formato {'indice': i, 'motivo': palavra}
                
                for i in range(total_botoes):
                    if _stop_flag: return
                    
                    sucesso_leitura = False
                    tentativas_waf = 0
                    
                    # Loop de retentativa para o MESMO documento caso o Firewall bloqueie
                    while not sucesso_leitura and tentativas_waf < 3:
                        if _stop_flag: return
                        try:   
                            botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                            if i >= len(botoes_atuais): break

                            botao_foco = botoes_atuais[i]
                            pular_este_arquivo = False
                            motivo_reprovacao = "Desconhecido" 
                            bateu_no_waf = False

                            if tentativas_waf == 0:
                                print(f"Mapeando documento {i+1} de {total_botoes}...")
                            else:
                                print(f"Retentando documento {i+1} (Tentativa WAF {tentativas_waf + 1}/3)...")

                            # Clica para abrir a aba de "Visualizar"
                            container_pai = botao_foco.find_element(By.XPATH, './../..')
                            botao_visualizar = container_pai.find_element(By.XPATH, ".//a[contains(normalize-space(.), 'Visualizar Instrumento Coletivo')]")

                            janelas_antes = set(browser.window_handles)
                            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_visualizar)
                            time.sleep(1)
                            browser.execute_script("arguments[0].click();", botao_visualizar)

                            WebDriverWait(browser, 15).until(EC.number_of_windows_to_be(len(janelas_antes) + 1))
                            nova_aba = (set(browser.window_handles) - janelas_antes).pop()
                            browser.switch_to.window(nova_aba)
                    
                            texto_da_pagina_lower = ""
                            tempo_limite = 30 
                            tamanho_anterior = 0
                            contador_estabilidade = 0
                            
                            for _ in range(tempo_limite):
                                if _stop_flag: break
                                try:
                                    texto_cru = browser.find_element(By.TAG_NAME, "body").text
                                    texto_cru_lower = texto_cru.lower()
                                    
                                    # DETECÇÃO DO FIREWALL (WAF)
                                    # WAF = Web Application Firewall (forma do site web se "proteger" contra bots, no caso do mediador ele mostra a preview do documento com apenas o texto "requested url was rejected..." ao invés da CCT de verdade )
                                    if "requested url was rejected" in texto_cru_lower or "your support id is" in texto_cru_lower:
                                        bateu_no_waf = True
                                        break # Sai do loop de espera imediatamente
                                    
                                    texto_cru_limpo = remove_acentos(texto_cru_lower)
                                    texto_temp = re.sub(r'\s+', ' ', texto_cru_limpo)
                                    
                                    tamanho_atual = len(texto_temp.strip())
                                    
                                    if tamanho_atual < 100 or "ocorreu um erro" in texto_temp or "service unavailable" in texto_temp:
                                        time.sleep(1)
                                        continue
                                    
                                    # Lógica de Estabilidade
                                    if tamanho_atual == tamanho_anterior:
                                        contador_estabilidade += 1
                                    else:
                                        contador_estabilidade = 0 
                                        tamanho_anterior = tamanho_atual
                                    
                                    if contador_estabilidade >= 2:
                                        texto_da_pagina_lower = texto_temp
                                        break # Sai do loop com o texto 100% completo
                                        
                                except:
                                    pass 
                                    
                                time.sleep(1) 
                            
                            # TRATAMENTO PÓS-ESPERA (Decide o que fazer se bateu no WAF)
                            if bateu_no_waf:
                                tempo_espera = 30 * (tentativas_waf + 1)
                                print(f"[WAF] Firewall do governo detectado! Resfriando IP por {tempo_espera} segundos...")
                                browser.close() 
                                browser.switch_to.window(original_window) 
                                time.sleep(tempo_espera) 
                                tentativas_waf += 1
                                continue # O 'continue' joga o código pro início do 'while', retentando o mesmo arquivo! (espera um tempo e tenta baixar o MESMO arquivo novamente)
                                
                            # Se chegou até aqui sem bater no WAF, marca sucesso para sair do 'while'
                            sucesso_leitura = True
                            
                            # Trava de segurança: Se deu timeout normal (sem texto e sem WAF)
                            if len(texto_da_pagina_lower.strip()) < 100:
                                print("AVISO: Documento não carregou ou site deu erro. Reprovado por segurança.")
                                pular_este_arquivo = True
                                motivo_reprovacao = "Falha no carregamento da página (Timeout ou Vazia)"

                            # IDEIA: mover palavras_filtro e excecoes_lojas para um arquivo JSON ao lado do executável, assim dá pra editar as listas sem precisar gerar um novo .exe a cada mudança
                            palavras_filtro = [
                                "rodoviária", "rodov", "peix", "peixarias", "rodoviarias", "farmacêuticos",

                                "11.564.609/0001-00", "90.818.667/0001-99", "90.093.345/0001-20", "00.388.455/0001-26",
                                "10.401.977/0001-74", "78.508.934/0001-26", "01.221.950/0001-09", "91.109.975/0001-08",
                                "90.813.726/0001-36", "92.963.693/0001-36", "92.963.297/0001-09", "92.963.651/0001-03",
                                "03.665.508/0001-05", "92.963.677/0001-43", "54.207.766/0001-70", "92.961.093/0001-39",
                                "93.074.383/0001-23", "97.056.840/0001-08", "91.167.858/0001-09", "84.307.917/0001-23", 
                                "76.687.615/0001-08", "59.842.294/0001-41", "76.682.236/0001-17", "15.106.815/0001-27",
                                "79.370.367/0001-57", "80.869.894/0001-90", "78.687.431/0001-65", "78.686.888/0001-55",
                                "77.841.682/0001-90", "04.243.203/0001-60", "79.371.423/0001-78", "82.662.735/0001-45",
                                "10.695.042/0001-48", "92.961.523/0001-12", "03.042.025/0001-46", "92.963.875/0001-07",
                                "89.948.905/0001-00", "92.685.460/0001-19", "91.563.114/0001-03", "10.400.614/0001-14",
                                "92.941.533/0001-96", "78.492.931/0001-41", "95.619.649/0001-00", "76.683.002/0001-94",
                                "76.682.210/0001-79", "80.623.622/0001-05", "80.920.085/0001-65", "76.683.010/0001-30",
                                "01.819.587/0001-28", "76.683.028/0001-32", "10.807.572/0001-30", "78.206.786/0001-95",
                                "82.662.750/0001-93", "06.037.567/0001-19", "79.266.730/0001-99", "76.618.222/0001-34",
                                "80.292.634/0001-02", "78.636.065/0001-15", "78.972.650/0001-96", "49.087.273/0001-04",
                                "92.517.457/0001-96", "03.499.644/0001-64", "43.450.014/0001-10", "52.806.460/0001-05",
                                "60.746.419/0001-19", "62.809.777/0001-59", "62.660.410/0001-16", "38.891.073/0001-93",
                                "62.202.759/0001-04", "62.803.085/0001-01", "15.336.300/0001-13", "62.216.627/0001-31",
                                "60.976.404/0001-47", "13.501.619/0001-21", "82.512.864/0001-57", "10.834.792/0001-53",
                                "83.797.191/0001-91", "08.580.010/0001-55"
                            ]
                             
                            # 11.564.609/0001-00 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS DO VALE DO SINOS - SINDIGENEROS/VALE
                            # 90.818.667/0001-99 - SINDICATO INTERMUNICIPAL DO COM. VAREJ. DE GEN. ALIM.
                            # 01.221.950/0001-09 - FEDERACAO NACIONAL DOS CONCESS E DISTRS DE VEICULOS
                            # 90.093.345/0001-20 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS
                            # 10.401.977/0001-74 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS DE CACHOEIRA DO SUL
                            # 90.813.726/0001-36 - SINDICATO DO COM ATAC DE AL E BEB EM GERAL NO EST RS
                            # 92.963.693/0001-36 - SIND COM ATAC DE PRODUTOS QUIM P IND LAV E DROG MED P A
                            # 92.963.297/0001-09 - SIND COM ATAC DE GENEROS ALIMENTICIOS DE PORTO ALEGRE
                            # 92.963.651/0001-03 - SIND. INTERM. COM. ATAC. MAT.CONSTR. LOUCAS, TINTAS, FERRAG.
                            # 03.665.508/0001-05 - SINDICATO DO COMERCIO ATACADISTA DO ESTADO DO RIO GRANDE DO
                            # 92.963.677/0001-43 - SIND COM ATAC DE TECIDOS VESTUARIO ERe ARMARINHO DE PA
                            # 54.207.766/0001-70 - SINDICATO NACIONAL DO COMERCIO TRANSPORTADOR-REVENDEDOR-RETAL DE COMBUSTIVEIS
                            # 92.961.093/0001-39 - SIND TRAB COM MINERIOS DERIV PETROLEO NO EST RGS
                            # 93.074.383/0001-23 - SINDICATO EMPREGADOS AGENTES AUTONOMOS COMERC ESTADO RS 
                            # 97.056.840/0001-08 - SINDICATO DAS EMPRESAS DE GARAGENS, ESTACIONAMENTO E DE LIMPEZA E CONSERVACAO DE VEICULOS NO ESTADO DO RIO GRANDE DO SUL
                            # 91.167.858/0001-09 - SINDICATO DOS REPRESENTANTES COMERCIAIS DE LAJEADO
                            # 00.388.455/0001-26 - SIND DO COMERCIO VAREJ DE GENEROS ALIM DE FARROUPILHA
                            # 91.109.975/0001-08 - SIND. DO COM. VAREJISTA DE GENEROS ALIM. DE CX. DO SUL 
                            # 84.307.917/0001-23 - SIND DOS TRAB COM MIN E DERIV DE PETROLEO NO EST S CAT
                            # 59.842.294/0001-41 - SIND NACIONAL EMPR DISTRIBUIDORAS PRODUTOS SIDERURGICOS
                            # 15.106.815/0001-27 - SINDICATO DAS EMPRESAS CEREALISTAS DO ESTADO DO PARANA
                            # 79.370.367/0001-57 - SIND DO COM VAREJ DE PROD FARMACEUT, DE USO HUMANO E ANIMAL
                            # 80.869.894/0001-90 - SIND DOS MOTORISTAS, CONDUT. DE VEIC. RODOV URBANOS E EM GERAL, TRAB.TRANSP. ROD. PBCO
                            # 78.687.431/0001-65 - SINDICATO DOS MOTORISTAS, CONDUTORES DE VEICULOS RODOVIARIOS URBANOS E EM GERAL, TRABALHADORES EM TRANSPORTES RODOVIARIOS DE DOIS VIZINHOS - SINTRODOV
                            # 78.686.888/0001-55 - SIND DOS TRAB EM TRANSP RODOVIARIOS DE FRANC BELTRAO
                            # 77.841.682/0001-90 - SINDICATO DOS TRAB EM TRANSP RODOVIARIOS DE CASCAVEL PR
                            # 04.243.203/0001-60 - SINDICATO DOS CONCESSIONARIOS E DISTRIBUIDORES DE VEICULOS NO ESTADO DO RIO GRANDE DO SUL - SINCODIV/RS
                            # 79.371.423/0001-78 - SINDICATO DAS EMPRESAS DE SERVICOS CONTABEIS E DAS EMPRESAS DE ASSESSORAMENTO, PERICIAS, INFORMACOES E PESQUISAS DE BLUMENAU E REGIAO
                            # 82.662.735/0001-45 - SIND DO COM VAR DE PROD FARMAC DO VALE DO ITAJAI
                            # 10.695.042/0001-48 - SINDICATO DAS EMPRESAS DE SISTEMAS ELETRONICOS DE SEGURANCA DO ESTADO DE SANTA CATARINA - SIESE-SC
                            # 92.961.523/0001-12 - SINDICATO DO COMERCIO DE VEICULOS E DE PECAS E ACESSORIOS PA
                            # 03.042.025/0001-46 - SINDICATO DO COM VAR MAT OPTICO FOTO E CINE DO RS
                            # 92.963.875/0001-07 - SINDICATO DO COM VAREJ DE PROD FARMACEUTICOS NO EST RGS
                            # 89.948.905/0001-00 - SINDICATO DOS ESTABELECIMENTOS DE PRESTACAO DE SERVICOS FUNERARIOS DO ESTADO DO RIO GRANDE DO SUL
                            # 92.685.460/0001-19 - OCERGS-SINDICATO E ORGANIZACAO DAS COOPERATIVAS DO ESTADO DO RS
                            # 10.400.614/0001-14 - FEDERACAO INTERMUNICIPAL DE SINDICATOS DE TRABALHADORES NO COMERCIO DE BENS E SERVICO DA FORCA SINDICAL NO ESTADO DO RIO GRANDE DO SUL - FETRACOS/RS
                            # 92.941.533/0001-96 - SINDICATO DO COMERCIO ATACADISTA DE MADEIRAS DE PA
                            # 91.563.114/0001-03 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS E DO COMERCIO VAREJISTA DE SUPERMERCADOS E HIPERMERCADOS DE PELOTAS
                            # 78.492.931/0001-41 - SINCODIV -SINDICATO DOS CONCESSIONARIOS E DISTRIBUIDORES DE VEICULOS NO ESTADO DE SANTA CATARINA
                            # 95.619.649/0001-00 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS E DO
                            # 80.623.622/0001-05 - SIND DO COM VAREJ DE PRODUTOS FARMACEUTICOS DO OES CAT
                            # 76.683.002/0001-94 - SINDICATO DO COM ATAC DE DROGAS E MEDIC NO ESTADO DO PR
                            # 76.682.210/0001-79 - SINDICATO DO COM VAREJ DE PRODUTOS FARM DO ESTADO DO PR
                            # 76.687.615/0001-08 - SINDICATO DO COMERCIO ATACADISTA DE MADEIRAS DO PARANA
                            # 76.682.236/0001-17 - SINDICATO DO COM DE VEIC PECAS E ACES PARA VEIC, MOT E BIC, COM IMPORT, DISTR DE AUTOP E MOTOP, ROL E ACES NO EST DO PARANA
                            # 78.508.934/0001-26 - SINDICATO DO COM VAR DE GEN ALIM DO ALTO IRANI SC
                            # 80.920.085/0001-65 - SINDICATO DO COMERCIO VAREJISTA DE MATERIAL OPTICO, FOTOGRAF, CNPJ n. 80.920.085/0001-65
                            # 76.683.010/0001-30 - SINDICATO DO COMERCIO ATACADISTA E DISTRIBUIDORES DO ES"TADO
                            # 01.819.587/0001-28 - SINDICATO DOS CONCESSIONARIOS E DISTRIBUIDORES DE VEICU
                            # 76.683.028/0001-32 - SINDICATO INTERMUNICIPAL DO COMERCIO VAREJISTA DE MATERIAIS
                            # 10.807.572/0001-30 - SINDICATO DO COMERCIO ATACADISTA E DISTRIBUIDOR DE JOACABA E
                            # 78.206.786/0001-95 - SIND DOS ESTABELECIMENTOS DE SERV FUNERARIO EST PARANA
                            # 82.662.750/0001-93 - SINDICATO DO COMERCIO ATACADISTA DO VALE DO ITAJAI SINCAVI
                            # 06.037.567/0001-19 - SINDICATO DO COMERCIO ATACADISTA E DISTRIBUIDOR DO NORTE E N
                            # 79.266.730/0001-99 - SIND LOJ COM VAR GEN ALIM MAQ FER TINT MAT EL AP ELE UM
                            # 76.618.222/0001-34 - SINDICATO DOS LOJISTAS DO COM. E DO COM. VAREJ. DE MAQ. FERRAGENS TINTAS E MAT. ELETRICO E AP. ELETRODOM. DE CURITIBA E REGIAO METROPOLITANA
                            # 80.292.634/0001-02 - SIND COM VAR FER TIN MAD MAT ELET HID MAT CONS MGA REG
                            # 78.636.065/0001-15 - SINDICATO DO COMERCIO VAREJISTA DE PRODUTOS FARMACEUTICOS DE
                            # 78.972.650/0001-96 - SINDICATO DO COMERCIO VAREJISTA DE VEICULOS DE LONDRINA
                            # 49.087.273/0001-04 - SINCOVAGA SIND. DO COM.VAR. DE GEN.ALIM. DE MERCADOS ARM.MER
                            # 92.517.457/0001-96 - SINDICATO DO COM. VAREJ. DE GENEROS ALIMENTICIOS E DO COM. VAR. DE SUPERMERCADOS E HIPERMERCADOS DOS VALES DO RIO PARDO E TAQUARI - SINDIGENEROS
                            # 03.499.644/0001-64 - SICAP - SIND DO COMERC ATACADISTA, IMPORTADOR, EXPORTADOR E DISTRIB DE PECAS, ROLAMENTOS, ACESSORIOS E COMPONENTES PARA INDUST E PARA VEICULOS EST SP
                            # 43.450.014/0001-10 - SINDICATO DO COMERCIO ATACADISTA, IMPORTADOR E EXPORTADOR DE PRODUTOS QUIMICOS E PETROQUIMICOS NO ESTADO DE SAO PAULO
                            # 52.806.460/0001-05 - SINCAMESP SIND COMERCIO ATACADISTA IMPORTADOR EXPORTADOR E DIST DE DROGAS MEDICAMENTOS C PERFUMARIAS COSMETICOS E ART DE TOUCADOR NO ESTADO SP,
                            # 60.746.419/0001-19 - SINDICATO DO COMERCIO ATACADISTA, IMPORTADOR, EXPORTADOR E DISTRIBUIDOR DE COUROS, PELES E SINTETICOS NO ESTADO DE SAO PAULO SINACOUROS
                            # 62.809.777/0001-59 - SIND DO COMERCIO ATACADISTA DE LOUCAS TINTAS E FERRAGENS NO
                            # 62.660.410/0001-16 - SINDICATO DO COMERCIO ATACADISTA DE PAPEL, PAPELAO, ARTIGOS
                            # 38.891.073/0001-93 - SINDICATO DO COMERCIO ATACADISTA DE SUCATA FERROSA E NAO FERROSA DO ESTADO DE SAO PAULO-SP
                            # 62.202.759/0001-04 - SIND COMERCIO ATACADISTA DE TEC VEST E ARM DO EST DE SP
                            # 62.803.085/0001-01 - SIND COM ATAC VIDRO PLANOCRISTAIS ESPELHOS EST S PAULO
                            # 15.336.300/0001-13 - SINAPESP - SINDICATO PATRONAL DAS EMPRESAS DE APARAS DE PAPEL E PAPELAO DO ESTADO DE SAO PAULO
                            # 62.216.627/0001-31 - SINDICATO DO COMERCIO VAREJISTA DOS FEIRANTES DE SAO PAULO - SINCOFER
                            # 60.976.404/0001-47 - SIN EMPREG AG AUT COM EMP ASS P IF PQ EMP SERV CONT ESP
                            # 13.501.619/0001-21 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS DE S
                            # 82.512.864/0001-57 - ORGANIZACAO DAS COOPERATIVAS DO ESTADO DE STA CATARINA
                            # 10.834.792/0001-53 - SINDICATO DO COMERCIO VAREJISTA DE GENEROS ALIMENTICIOS DE PASSO FUNDO
                            # 83.797.191/0001-91 - SINDICATO DAS EMPRESAS DE SERVICOS CONTABEIS, ASSESSORAMENTO, PERICIAS, INFORMACOES E PESQUISAS NO ESTADO DE SANTA CATARINA - SESCON/SC 
                            # 08.580.010/0001-55 - SESCON SUL DE SANTA CATARINA - SINDICATO DAS EMPRESAS DE SERV.CONT.E DAS EMPR.DE ASSESS. PERICIA, INFORM.E PESQ. DO SUL DE SC

                            excecoes_lojas = {
                                430: ["79.266.730/0001-99"],
                                437: ["79.266.730/0001-99"],
                                467: ["78.508.934/0001-26"],
                                478: ["76.683.010/0001-30", "76.687.615/0001-08", "76.683.002/0001-94", "76.682.210/0001-79"],
                                591: ["76.683.010/0001-30", "76.687.615/0001-08", "76.683.002/0001-94"],
                                610: ["76.683.010/0001-30", "76.687.615/0001-08", "76.683.002/0001-94", "76.682.210/0001-79"],
                            }
                            cnpjs_ignorados = [
                                remove_acentos(c.lower()) for c in excecoes_lojas.get(int(numloja), [])
                            ]
                            
                            excecao_encontrada = any(cnpj in texto_da_pagina_lower for cnpj in cnpjs_ignorados) #percorre a lista de cnpjs

                            if excecao_encontrada:
                                print(f"Exceção encontrada na loja {numloja} — documento aprovado, filtro ignorado.")
                                pular_este_arquivo = False
                                motivo_reprovacao = None
                            else: # se for false
                                for palavra in palavras_filtro: # roda o filtro normal
                                    palavra_limpa_filtro = remove_acentos(palavra.lower())
                                    if palavra_limpa_filtro in texto_da_pagina_lower:
                                        print(f"Documento REPROVADO no filtro. Palavra '{palavra}' encontrada.")
                                        pular_este_arquivo = True
                                        motivo_reprovacao = f"Palavra '{palavra}'"
                                        break

                            if not pular_este_arquivo:
                                print(f"Documento APROVADO! Adicionado à fila de download principal.")
                                indices_aprovados.append(i) 
                            else:
                                documentos_reprovados.append({'indice': i, 'motivo': motivo_reprovacao})

                            browser.close() 
                            browser.switch_to.window(original_window) 
                            time.sleep(1)

                        except Exception as e_map:
                            print(f"Erro ao mapear o arquivo {i+1}: {e_map}")
                            log_erro(arquivo_log, f"Erro ao mapear o documento {i+1}, ele NÃO foi baixado: {str(e_map)}", cnpj, numloja)
                            browser.switch_to.window(original_window) 
                            break # Sai do while em caso de erro crítico que não seja o WAF

                qtd_aprovados = len(indices_aprovados)
                
                if qtd_aprovados == 0:
                    print("Nenhum documento da lista passou no filtro de palavras.")
                    registrar_sem_doc(numloja, cnpj, municipio, codsind)
                    
                else:
                    print(f"Mapeamento concluído. {qtd_aprovados} documento(s) serão baixados.")
                    
                    usar_pasta_municipio_na_chamada = qtd_aprovados > 1

                    for contador_baixados, index_real in enumerate(indices_aprovados, start=1):
                        if _stop_flag: return
                        
                        try:
                            janelas_antes_download = set(browser.window_handles)
                            botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                            
                            if index_real >= len(botoes_atuais): break
                                
                            botao_clique_real = botoes_atuais[index_real]

                            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_clique_real)
                            time.sleep(1.5)
                            arquivos_antes = set(glob.glob(os.path.join(download_dir, "*.doc"))) # "foto" da pasta antes do clique
                            browser.execute_script("arguments[0].click();", botao_clique_real)

                            print(f"Baixando documento ({contador_baixados}/{qtd_aprovados}) ...")
                            time.sleep(5) # Respiro do servidor para gerar o arquivo
                           
                            try:
                                WebDriverWait(browser, 5).until(lambda drv: len(drv.window_handles) > len(janelas_antes_download)) # Limpeza de janelas extras de download
                                novas_janelas = set(browser.window_handles) - janelas_antes_download
                                for aba in novas_janelas:
                                    browser.switch_to.window(aba)
                                    browser.close()
                                browser.switch_to.window(original_window)
                            except:
                                pass

                            sucesso, pdf_path, folder_path = renomear_arquivo_baixado(
                                codsind, 
                                municipio, 
                                cidade_sind, 
                                cnpj,
                                f"{numloja}__{contador_baixados}" if usar_pasta_municipio_na_chamada else numloja,
                                arquivos_antes,
                                usar_pasta_municipio=usar_pasta_municipio_na_chamada
                            )
                            
                            if sucesso:
                                
                                # Move para pasta de atrasados se necessário
                                if dias_de_atraso < 0:
                                    print(f"Documento está atrasado ({abs(dias_de_atraso)} dias). Movendo para ATRASADOS...")
                                    
                                    if usar_pasta_municipio_na_chamada and folder_path:
                                        # Cria a pasta de destino se não existir
                                        destino_atrasado_pasta = os.path.join(atrasados_dir, os.path.basename(folder_path))
                                        os.makedirs(destino_atrasado_pasta, exist_ok=True)
                                     
                                        if pdf_path and os.path.exists(pdf_path):
                                            destino_pdf = os.path.join(destino_atrasado_pasta, os.path.basename(pdf_path))
                                            if os.path.exists(destino_pdf):
                                                os.remove(destino_pdf)
                                            shutil.move(pdf_path, destino_pdf)
                                            
                                        try:
                                            os.rmdir(folder_path) # Limpa a pasta vazia que sobrou na raiz do Downloads
                                        except:
                                            pass

                                    elif pdf_path:
                                        destino_atrasado = os.path.join(atrasados_dir, os.path.basename(pdf_path))   # Movimentação normal para arquivos soltos (sem pasta)
                                        if os.path.exists(destino_atrasado):
                                            os.remove(destino_atrasado)
                                        shutil.move(pdf_path, atrasados_dir)

                        except Exception as e_down:
                            import traceback
                            print(traceback.format_exc())
                            log_erro(arquivo_log, f"Erro ao executar o download do arquivo aprovado {contador_baixados}: {str(e_down)}", cnpj, numloja)

                # Fase de auditoria (DEBUG MODE para baixar arquivos filtrados)
                if MODO_DEBUG_FILTRO and documentos_reprovados:
                    print(f"\n[MODO DEBUG] Iniciando download de {len(documentos_reprovados)} documento(s) REPROVADOS para auditoria...")
                    
                    usar_pasta_debug = len(documentos_reprovados) > 1

                    for contador_debug, doc_reprovado in enumerate(documentos_reprovados, start=1):
                        if _stop_flag: return
                        index_real = doc_reprovado['indice']
                        motivo = doc_reprovado['motivo']
                        
                        try:    
                            janelas_antes_download = set(browser.window_handles)
                            botoes_atuais = browser.find_elements(By.XPATH, '//*[@id="btnDownload"]/a')
                            
                            if index_real >= len(botoes_atuais): break
                                
                            botao_clique_real = botoes_atuais[index_real]
                            browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_clique_real)
                            time.sleep(1.5)
                            arquivos_antes = set(glob.glob(os.path.join(download_dir, "*.doc"))) # "foto" da pasta antes do clique
                            browser.execute_script("arguments[0].click();", botao_clique_real)

                            print(f"Baixando documento REPROVADO ({contador_debug}/{len(documentos_reprovados)}) ...")
                            time.sleep(5) # Respiro do servidor
                           
                            try:
                                WebDriverWait(browser, 5).until(lambda drv: len(drv.window_handles) > len(janelas_antes_download)) # Limpeza de janelas extras de download
                                novas_janelas = set(browser.window_handles) - janelas_antes_download
                                for aba in novas_janelas:
                                    browser.switch_to.window(aba)
                                    browser.close()
                                browser.switch_to.window(original_window)
                            except:
                                pass

                            sucesso, pdf_path, folder_path = renomear_arquivo_baixado(  # Renomeia, Converte e ROTEIA direto para a pasta DEBUG
                                codsind, 
                                municipio, 
                                cidade_sind, 
                                cnpj,
                                f"{numloja}___filtrado_{contador_debug}" if usar_pasta_debug else f"{numloja}__filtrado",
                                arquivos_antes,
                                usar_pasta_municipio=usar_pasta_debug,
                                is_debug=True 
                            )
                            
                            if sucesso:
                                nome_arquivo_salvo = os.path.basename(pdf_path) if pdf_path else "Desconhecido"

                                linha_log = f"[{datetime.now().strftime('%d/%m/%Y %H:%M')}] Loja: {numloja.zfill(3)} | Sindicato: {codsind.zfill(4)} | Arquivo Salvo: {nome_arquivo_salvo} | Motivo: {motivo}\n"
                                with open(caminho_relatorio_filtros, "a", encoding="utf-8") as f:
                                    f.write(linha_log)
                                    
                            else:
                                print("Não foi possível processar o arquivo reprovado.")

                        except Exception as e_debug:
                            import traceback
                            print(traceback.format_exc())
                            log_erro(arquivo_log, f"Erro ao executar o download do arquivo de auditoria {contador_debug}: {str(e_debug)}", cnpj, numloja)
                        
                    with open(caminho_relatorio_filtros, "a", encoding="utf-8") as f:
                        f.write("\n")

            browser.get('http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo')
            if _stop_flag: return

    except Exception as e: # Mudança aqui
        mensagem_erro = str(e)

        print(f"Erro na função movimentacao_site: {mensagem_erro}")

        contagem_anterior = _falhas_consecutivas.get(numloja, 0)
        contagem_atual = contagem_anterior + 1
        _falhas_consecutivas[numloja] = contagem_atual

        if contagem_atual >= 2:
            log_erro(arquivo_log, f"Erro persistente (após reinício do navegador): {mensagem_erro}", cnpj, numloja)
        else:
            print(f"Erro detectado (tentativa {contagem_atual}). Reiniciando navegador antes de tentar novamente...")

        return "SESSION_DEAD"

#############################################################################
def criar_browser(download_dir):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    #chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu") 
    chrome_options.add_argument("--window-position=-2400,-2400")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-features=CalculateNativeWinOcclusion")

    final_download_path = os.path.abspath(download_dir)
    print(f"INFO: Configurando diretório de download do Chrome para: {final_download_path}")
    prefs = {
        "download.prompt_for_download": False, 
        "download.directory_upgrade": True,
        "profile.default_content_settings.popups": 0,
        "download.default_directory": final_download_path,
        "safebrowsing.enabled": False, 
        "safebrowsing.disable_download_protection": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=chrome_options)
    driver.get('http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo')
    return driver
#############################################################################

def limpar_diretorio(diretorio):
    # Apaga só as pastas que o próprio robô cria, o resto da pasta de trabalho fica intacto (dessa forma podemos manter um log permanente de todas as rotinas)
    pastas_do_robo = ["Downloads", "Histórico", "Logs"]
    for nome in pastas_do_robo:
        caminho = os.path.join(diretorio, nome)
        if os.path.isdir(caminho):
            try:
                shutil.rmtree(caminho)
                print(f"[LIMPEZA] Pasta removida: {nome}")
            except Exception as e:
                print(f"[LIMPEZA] Erro ao remover '{nome}': {e}")

#############################################################################

def main(caminho_excel=None, base_path_param=None, loja_inicial=None, update_callback=None): #  parte principal do código
    global _stop_flag, _browser_driver
    global mediador_dir, historico_dir, download_dir, atrasados_dir
    global pasta_log, arquivo_log, diretorio_log, caminho_log_hist

    # MODO DEBUG DOS FILTROS (baixa o que foi filtrado para auditoria)
    global MODO_DEBUG_FILTRO, filtrados_debug_dir, caminho_relatorio_filtros
    # para testar o modo debug essa variável precisa receber 'True'
    MODO_DEBUG_FILTRO = False # Mude para False quando for criar o executável de produção

    _stop_flag = False
    _browser_driver = None
    data_execucao = datetime.now()

    if base_path_param:
        print(f"Diretório de trabalho definido pela interface: {base_path_param}")
        mediador_dir = base_path_param
        historico_dir = os.path.join(mediador_dir, "Histórico")
        download_dir = os.path.join(mediador_dir, "Downloads")
        atrasados_dir = os.path.join(download_dir, "ATRASADOS")
        
        filtrados_debug_dir = os.path.join(download_dir, "FILTRADOS_DEBUG")  # Criação do diretório de DEBUG
        caminho_relatorio_filtros = os.path.join(filtrados_debug_dir, "Relatorio_Filtros_Reprovados.txt")
        
        pasta_log = os.path.join(mediador_dir, "Logs")
        timestamp_sessao = datetime.now().strftime("%d-%m-%Y_%H-%M")
        arquivo_log = os.path.join(pasta_log, f"Log Erros_{timestamp_sessao}.txt")
        diretorio_log = os.path.dirname(arquivo_log)
        caminho_log_hist = None
        #LOG DIRETORIO HISTÓRICO
        os.makedirs(diretorio_log, exist_ok=True)
        os.makedirs(download_dir, exist_ok=True)
        os.makedirs(atrasados_dir, exist_ok=True)
        os.makedirs(historico_dir, exist_ok=True)
        os.makedirs(pasta_log, exist_ok=True)
        organizar_Log_erro(PASTA_LOG_PERMANENTE) # mantém só os 30 logs mais recentes no Logs Permanentes

        if MODO_DEBUG_FILTRO:  # Se o modo debug estiver ativo, coloca o cabeçalho temporal no relatório
            os.makedirs(filtrados_debug_dir, exist_ok=True) # Garante que a pasta de filtrados existe
            with open(caminho_relatorio_filtros, "a", encoding="utf-8") as f:
                f.write(f"\n=== NOVA EXECUÇÃO DE DEBUG: {datetime.now().strftime('%d/%m/%Y às %H:%M')} ===\n")

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
           
            for row in df.itertuples(index=False):
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

    # Chama a função criar_browser    
    try:
        _browser_driver = criar_browser(download_dir)
        browser = _browser_driver
    except Exception as e:
        log_erro(arquivo_log, f"Ocorreu um erro ao abrir o Chrome: {str(e)}")
        return

    if browser is not None:
        print("Navegador Iniciado")
        try:
            browser.get('http://www3.mte.gov.br/sistemas/mediador/ConsultarInstColetivo')
            print("Página carregada")

            contador_lojas_processadas = 0        
            RESTART_A_CADA = 50

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
                
                # padroniza para evitar o pandas trasnformar em decimal
                numloja_atual = str(int(row[0]))
                if not iniciar_processamento_agora:
                    if numloja_atual == str(loja_final_para_iniciar):
                        print(f"--- Loja alvo '{loja_final_para_iniciar}' encontrada. Iniciando processamento. ---")
                        iniciar_processamento_agora = True
                    else:
                        print(f"Pulando loja: {numloja_atual}...")
                        continue
                
                if numloja_atual == 'None': continue
                
                # indices da planilha acompanhamentos
                cnpj = str(row[2])
                uf = str(row[4])
                municipio = str(row[5])
                cidade_sind = str(row[9])
                codsind = str(row[1])
                numloja = numloja_atual

                sind_bruto = codsind.strip()
                if sind_bruto.upper() == "FECHADA":
                    registrar_loja_pulada(numloja, cnpj, municipio, codsind, "Loja fechada - foi pulada")
                    print(f"[STATUS] Loja {numloja} fechada (CÓD. SIND='{codsind}'). Pulada.")
                    continue
                if not sind_bruto or sind_bruto.lower() in ("nan", "none"):
                    registrar_loja_pulada(numloja, cnpj, municipio, "não informado", "Código do sindicato não informado - foi pulada")
                    print(f"[STATUS] Loja {numloja} sem código de sindicato. Pulada.")
                    continue 

                dias_de_atraso = 0
                data_base_str = "Desconhecida"
                
                if len(row) > 7 and pd.notna(row[7]):
                    try:
                        valor_celula = row[7] # coluna H da planilha acompanhamentos tem a data-base (oitava coluna = row 7 pq a contagem começa do zero, 0 1 2 3 4 5 6 7)
                        
                        if isinstance(valor_celula, (int, float)):
                            data_base = pd.to_datetime(valor_celula, unit='D', origin='1899-12-30')
                        else:
                            data_base = pd.to_datetime(valor_celula, dayfirst=True)
                            
                        data_hoje = datetime.now()
                        
                        data_base = data_base.replace(hour=0, minute=0, second=0, microsecond=0)
                        data_hoje = data_hoje.replace(hour=0, minute=0, second=0, microsecond=0)
                        
                        diferenca = data_base - data_hoje
                        dias_de_atraso = diferenca.days
                        data_base_str = data_base.strftime('%d/%m/%Y')
                        
                    except Exception as e_data:
                        print(f"Aviso: Não foi possível calcular a data da Coluna H para a Loja {numloja}. Erro: {e_data}")
                        dias_de_atraso = 0
                        
                primeira_letra_municipio = primeira_letra(municipio) 

                if update_callback:
                    update_callback(f"{numloja} | CNPJ: {cnpj}")
                
                # PRINTS AGRUPADOS E FÁCEIS DE LER
                print("\n" + "="*70)
                print(f"Processando: Loja {numloja} | CNPJ {cnpj} | Código Sindicato {codsind} | Município {municipio}")
                
                if dias_de_atraso < 0:
                    print(f"[STATUS] ATRASADO em {abs(dias_de_atraso)} dias! (Data-base lida: {data_base_str})")
                else:
                    print(f"[STATUS] DENTRO DO PRAZO. Faltam {dias_de_atraso} dias. (Data-base lida: {data_base_str})")
                print("="*70 + "\n")

                if not cnpj or cnpj == 'None':
                    print(f"CNPJ ausente ou inválido, pulando: CNPJ lido: {cnpj}")
                    continue

                if contador_lojas_processadas > 0 and contador_lojas_processadas % RESTART_A_CADA == 0:  # Retart periódico do driver
                    print(f"\n>>> {contador_lojas_processadas} lojas processadas. Reiniciando o navegador para liberar memória... <<<\n")
                    try:
                        browser.quit()
                    except Exception as e_quit:
                        print(f"Aviso ao fechar navegador antigo: {e_quit}")

                    time.sleep(2)

                    try:
                        _browser_driver = criar_browser(download_dir)
                        browser = _browser_driver
                    except Exception as e_restart:
                        # Sem navegador não tem como continuar, encerra a execução (o finally lá embaixo ainda roda)
                        log_erro(arquivo_log, f"Erro crítico ao reiniciar o navegador, execução encerrada: {str(e_restart)}")
                        return

                MAX_TENTATIVAS_LOJA = 3
                tentativa_loja = 0
                processado = False

                while not processado and tentativa_loja < MAX_TENTATIVAS_LOJA:
                    resultado = movimentacao_site(cnpj, browser, uf, municipio, primeira_letra_municipio, codsind, numloja, cidade_sind, dias_de_atraso)

                    if resultado == "SESSION_DEAD":
                        print(f"\n>>> Sessão do navegador perdida. reiniciando driver para retomar na loja {numloja}... <<<\n")
                        try:
                            browser.quit()
                        except Exception as e_quit:
                            print(f"Aviso ao fechar navegador antigo: {e_quit}")
                        time.sleep(3)
                        try:
                            _browser_driver = criar_browser(download_dir)
                            browser = _browser_driver
                        except Exception as e_restart:
                            # Sem navegador não tem como continuar, encerra a execução (o finally lá embaixo ainda roda)
                            log_erro(arquivo_log, f"Erro crítico ao reiniciar o navegador, execução encerrada: {str(e_restart)}")
                            return
                        tentativa_loja += 1
                    else:
                        processado = True

                # Se a loja deu erro nas 3 tentativas, registra no histórico pra ela não sumir sem deixar rastro
                if not processado:
                    registrar_loja_pulada(numloja, cnpj, municipio, codsind, f"Erro após {MAX_TENTATIVAS_LOJA} tentativas - foi pulada")

                contador_lojas_processadas += 1
                _falhas_consecutivas.pop(numloja, None)

            if not _stop_flag:
                print("Processamento concluído.")
                email_enviado = False
                arquivo_log_permanente = os.path.join(PASTA_LOG_PERMANENTE, os.path.basename(arquivo_log))
                try:
                    email_enviado = enviar_relatorio(mediador_dir, data_execucao, arquivo_log_permanente, caminho_excel)  # caminho_excel: base de dados mostrada no email
                except Exception as e_email:
                    log_erro(arquivo_log, f"Erro ao enviar relatório por email: {e_email}")

                if email_enviado:
                    limpar_diretorio(mediador_dir)
                else:
                    print("[LIMPEZA] Email não enviado. Diretório NÃO limpo.")

        except Exception as e:
            log_erro(arquivo_log, f"Ocorreu um erro na execução principal: {str(e)}", cnpj=locals().get('cnpj'))
        finally:
            if browser:
                browser.quit()