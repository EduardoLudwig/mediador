import smtplib
import os
import glob
import subprocess
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
SEVENZIP_PATH = r"C:\Program Files\7-Zip\7z.exe"
LIMITE_ANEXO_MB = 18
TIMEOUT_SMTP = 180
TIMEOUT_7Z = 1800

################################################################

def carregar_config_email():
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(f"Arquivo de configuração não encontrado: {CONFIG_PATH}")

    config.read(CONFIG_PATH, encoding="utf-8")

    if "Email" not in config:
        raise KeyError("Seção [Email] não encontrada no config.ini. Adicione EmailRemetente, SenhaApp e EmailDestinatario.")

    email_cfg = config["Email"]
    remetente = email_cfg.get("EmailRemetente", "")
    senha = email_cfg.get("SenhaApp", "")
    destinatario = email_cfg.get("EmailDestinatario", "")

    if not remetente or not senha or not destinatario:
        raise ValueError("Configuração de email incompleta. Verifique EmailRemetente, SenhaApp e EmailDestinatario no config.ini")

    return remetente, senha, destinatario

################################################################

def contar_arquivos_download(diretorio_downloads):
    padrao = os.path.join(diretorio_downloads, "**", "*.pdf")
    arquivos = glob.glob(padrao, recursive=True)
    return len(arquivos)

################################################################

def compactar_pasta_7z(pasta_origem):
    pasta_origem = os.path.normpath(pasta_origem)

    if not os.path.isfile(SEVENZIP_PATH):
        raise FileNotFoundError(
            f"7-Zip não encontrado em: {SEVENZIP_PATH}\n"
            "Instale o 7-Zip: https://www.7-zip.org/"
        )

    data_atual = datetime.now()
    nome_arquivo = f"leitura_mediador_{data_atual.strftime('%d-%m-%Y')}.7z"
    pasta_pai = os.path.dirname(pasta_origem)
    caminho_7z = os.path.join(pasta_pai, nome_arquivo)

    try:
        resultado = subprocess.run(
            [
                SEVENZIP_PATH,
                "a", "-t7z", "-mx=5",
                caminho_7z,
                pasta_origem
            ],
            capture_output=True,
            text=True,
            timeout=TIMEOUT_7Z,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        if resultado.returncode != 0:
            _limpar_arquivo(caminho_7z)
            raise RuntimeError(f"Erro ao compactar com 7-Zip:\n{resultado.stderr}")

        return caminho_7z

    except subprocess.TimeoutExpired:
        _limpar_arquivo(caminho_7z)
        raise TimeoutError(f"7-Zip excedeu o tempo limite de {TIMEOUT_7Z}s na compactação")
    except Exception:
        _limpar_arquivo(caminho_7z)
        raise

################################################################

def enviar_relatorio(mediador_dir):
    try:
        remetente, senha, destinatario = carregar_config_email() # devolve tupla de 3 elementos
    except (FileNotFoundError, KeyError, ValueError) as e: # cobre cenários de falha
        print(f"[EMAIL] Erro ao carregar configuração: {e}")
        return False

    diretorio_downloads = os.path.join(mediador_dir, "Downloads")
    if os.path.isdir(diretorio_downloads):
        qtd_arquivos = contar_arquivos_download(diretorio_downloads)
    else:
        print(f"[EMAIL] Pasta Downloads não encontrada, contagem de arquivos será 0.")
        qtd_arquivos = 0
    data_atual = datetime.now()
    data_formatada = data_atual.strftime("%d/%m/%Y")

    print("[EMAIL] Compactando pasta de leitura...")
    try:
        caminho_7z = compactar_pasta_7z(mediador_dir)
    except Exception as e:
        print(f"[EMAIL] Erro ao compactar: {e}")
        return False

    print("[EMAIL] Preparando envio do email...")

    tamanho_mb = os.path.getsize(caminho_7z) / (1024 * 1024)
    if tamanho_mb > LIMITE_ANEXO_MB:
        print(f"[EMAIL] Arquivo de {tamanho_mb:.1f} MB excede o limite de {LIMITE_ANEXO_MB} MB. Email não enviado.")
        _limpar_arquivo(caminho_7z)
        return False

    msg = MIMEMultipart() # mensagem com corpo de texto + anexo. sem multipart teria só um dos dois
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = f"Relatório Robô Mediador - {data_formatada}"

    corpo = (
        "Olá,\n\n"
        "Segue em anexo a leitura completa do Robô Mediador.\n\n"
        f"Data da leitura: {data_formatada}\n"
        f"Quantidade de arquivos baixados: {qtd_arquivos}\n\n"
        "Att,\nRobô Mediador"
    )

    msg.attach(MIMEText(corpo, "plain", "utf-8"))

    try:
        with open(caminho_7z, "rb") as f:
            parte = MIMEBase("application", "x-7z-compressed")
            parte.set_payload(f.read())
            encoders.encode_base64(parte)
            nome_arquivo = os.path.basename(caminho_7z)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename=\"{nome_arquivo}\""
            )
            msg.attach(parte)
    except Exception as e:
        print(f"[EMAIL] Erro ao ler arquivo 7z para anexo: {e}")
        _limpar_arquivo(caminho_7z)
        return False

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=TIMEOUT_SMTP) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatario, msg.as_string())
        print(f"[EMAIL] Email enviado com sucesso para {destinatario}")
        _limpar_arquivo(caminho_7z)
        return True
    except smtplib.SMTPAuthenticationError:
        print("[EMAIL] Erro de autenticação. Verifique o email e a senha de app no config.ini")
        _limpar_arquivo(caminho_7z)
        return False
    except Exception as e:
        print(f"[EMAIL] Erro ao enviar email: {e}")
        _limpar_arquivo(caminho_7z)
        return False

################################################################

def _limpar_arquivo(caminho):
    try:
        if os.path.exists(caminho):
            os.remove(caminho)
            print(f"[EMAIL] Arquivo temporário removido: {caminho}")
    except Exception as e:
        print(f"[EMAIL] Aviso: não foi possível remover o arquivo temporário: {e}")
