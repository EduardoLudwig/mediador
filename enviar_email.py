import smtplib
import os
import html  # html.escape: evita que caracteres como & e < no caminho da base quebrem o email
import glob
import subprocess
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# Mesmo config.ini que o robô usa. Não dá pra usar a pasta do __file__ porque no .exe (--onefile) ela aponta pra uma pasta temporária que o PyInstaller cria ao abrir o programa, por isso funcionava rodando no vscode mas no .exe
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".RoboMediador", "config.ini")
SEVENZIP_PATH = r"C:\Program Files\7-Zip\7z.exe"
LIMITE_ANEXO_MB = 18
TIMEOUT_SMTP = 180
TIMEOUT_7Z = 1800

################################################################

def load_last_email_config():
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_PATH):
        return None

    config.read(CONFIG_PATH, encoding="utf-8")

    if "Email" not in config:
        return None

    email_cfg = config["Email"]
    remetente = email_cfg.get("EmailRemetente", "")
    senha = email_cfg.get("SenhaApp", "")
    destinatarios_raw = email_cfg.get("EmailDestinatario", "")

    destinatarios = [
        email.strip()
        for email in destinatarios_raw.replace(",", ";").split(";")
        if email.strip()
    ]

    return {
        "remetente": remetente,
        "senha": senha,
        "destinatarios": destinatarios,
    }

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
    destinatarios_raw = email_cfg.get("EmailDestinatario", "")

    if not remetente or not senha or not destinatarios_raw:
        raise ValueError("Configuração de email incompleta. Verifique EmailRemetente, SenhaApp e EmailDestinatario no config.ini")

    destinatarios = [
        email.strip()
        for email in destinatarios_raw.replace(",", ";").split(";")
        if email.strip()
    ]

    if not destinatarios:
        raise ValueError("EmailDestinatario vazio ou inválido no config.ini. Separe múltiplos emails por ponto e vírgula (;) ou vírgula (,).")

    return remetente, senha, destinatarios

################################################################

def contar_arquivos_download(diretorio_downloads):
    padrao = os.path.join(diretorio_downloads, "**", "*.pdf")
    arquivos = glob.glob(padrao, recursive=True)
    return len(arquivos)

################################################################

def compactar_pasta_7z(pasta_origem, data_execucao=None):
    pasta_origem = os.path.normpath(pasta_origem)

    if not os.path.isfile(SEVENZIP_PATH):
        raise FileNotFoundError(
            f"7-Zip não encontrado em: {SEVENZIP_PATH}\n"
            "Instale o 7-Zip: https://www.7-zip.org/"
        )

    data_ref = data_execucao if data_execucao else datetime.now()
    nome_arquivo = f"leitura_mediador_{data_ref.strftime('%d-%m-%Y')}.7z"
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

# O HTML do email é montado com TABELAS (e não <div>) porque o Outlook para Windows usa o motor do Word,
# que ignora largura, fundo e espaçamento em <div>. Tabelas funcionam no Outlook, no Gmail e no celular.
# A fonte é repetida em cada célula porque o Outlook não herda a fonte da tabela de fora.
FONTE = "font-family: Arial, sans-serif;"

def montar_corpo_html(qtd_arquivos, data_ref, caminho_base=None):
    # Corpo do email em HTML: base de dados usada e o resumo. A leitura completa vai no .7z em anexo.
    # Os estilos vão dentro de cada tag (style="...") porque os programas de email ignoram CSS separado
    verde = "#0F3B17"  # verde bem escuro, para diferenciar do email do robô CCTs ComuniQ

    # Base de dados (planilha Excel) que o robô leu. Só aparece se o caminho for informado
    base_html = ""
    if caminho_base:
        caminho_exibido = html.escape(os.path.normpath(caminho_base))  # normpath: mostra com '\' como no Windows
        base_html = f"""
      <div style="font-size: 13px; color: #666666;">Base de dados utilizada:</div>
      <p style="margin: 0 0 16px 0; font-size: 13px; color: #333333;">{caminho_exibido}</p>"""

    # width="600" é a largura que o Outlook respeita (não estica até a janela);
    # max-width: 100% deixa os outros programas e o celular encolherem quando a tela for menor
    return f"""
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width: 600px; max-width: 100%; border: 1px solid #dddddd;">
  <tr>
    <td bgcolor="{verde}" style="background-color: {verde}; padding: 18px 24px; {FONTE}">
      <div style="font-size: 20px; font-weight: bold; color: #ffffff;">Relatório do Robô Mediador</div>
      <div style="font-size: 14px; color: #E3F1E8; padding-top: 6px;">Download das Convenções Coletivas de Trabalho das lojas no Mediador</div>
      <div style="font-size: 14px; color: #E3F1E8;">{data_ref.strftime('%d/%m/%Y')}</div>
    </td>
  </tr>
  <tr>
    <td style="padding: 20px 24px; {FONTE} color: #333333;">{base_html}
      <p style="margin: 0 0 8px 0; font-size: 15px; color: #333333;">Olá,</p>
      <p style="margin: 0 0 12px 0; font-size: 15px; color: #333333;">Segue em anexo a leitura completa do Robô Mediador.</p>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="padding: 10px 0; border-bottom: 1px solid #eeeeee; {FONTE} font-size: 15px; font-weight: bold; color: #333333;">Arquivos baixados</td>
          <td width="60" align="right" style="padding: 10px 0; border-bottom: 1px solid #eeeeee; {FONTE} font-size: 22px; font-weight: bold; color: {verde};">{qtd_arquivos}</td>
        </tr>
      </table>
      <p style="margin: 24px 0 0 0; font-size: 15px; color: #333333;">Att,<br>Robô Mediador</p>
    </td>
  </tr>
</table>
"""

################################################################

def _log_erro_email(arquivo_log, mensagem):
    # anota erros no log permanente, serve para rastreabilidade
    data_log = datetime.now().strftime("%d-%m-%Y_%H-%M")
    layout_mensagem = (
        f"Erro na data: {data_log}\n"
        f"ENVIO DE EMAIL\n\n"
        f"{mensagem}\n"
        f"\n-------------------------------------------------\n\n"
    )
    print(layout_mensagem)
    if arquivo_log:
        with open(arquivo_log, "a", encoding="utf-8") as log:
            log.write(layout_mensagem)

################################################################

def enviar_relatorio(mediador_dir, data_execucao=None, arquivo_log=None, caminho_base=None):
    # caminho_base: planilha Excel usada na leitura; aparece no corpo do email (opcional)
    try:
        remetente, senha, destinatarios = carregar_config_email() # devolve tupla de 3 elementos
    except (FileNotFoundError, KeyError, ValueError) as e: # cobre cenários de falha
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao carregar configuração: {e}")
        return False

    pastas_obrigatorias = {
        "Downloads": os.path.join(mediador_dir, "Downloads"),
        "Histórico": os.path.join(mediador_dir, "Histórico"),
        "Logs": os.path.join(mediador_dir, "Logs"),
    }
    for nome_pasta, caminho_pasta in pastas_obrigatorias.items():
        if not os.path.isdir(caminho_pasta):
            _log_erro_email(arquivo_log, f"[EMAIL] Pasta obrigatória '{nome_pasta}' não encontrada em {caminho_pasta}. Email NÃO enviado.")
            return False

    diretorio_downloads = pastas_obrigatorias["Downloads"]
    qtd_arquivos = contar_arquivos_download(diretorio_downloads)
    data_ref = data_execucao if data_execucao else datetime.now()
    data_formatada = data_ref.strftime("%d/%m/%Y")

    print("[EMAIL] Compactando pasta de leitura...")
    try:
        caminho_7z = compactar_pasta_7z(mediador_dir, data_ref)
    except Exception as e:
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao compactar: {e}")
        return False

    print("[EMAIL] Preparando envio do email...")

    tamanho_mb = os.path.getsize(caminho_7z) / (1024 * 1024)
    if tamanho_mb > LIMITE_ANEXO_MB:
        _log_erro_email(arquivo_log, f"[EMAIL] Arquivo de {tamanho_mb:.1f} MB excede o limite de {LIMITE_ANEXO_MB} MB. Email não enviado.")
        _limpar_arquivo(caminho_7z)
        return False

    msg = MIMEMultipart() # mensagem com corpo de texto + anexo. sem multipart teria só um dos dois
    msg["From"] = remetente
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"Relatório Robô Mediador - {data_formatada}"

    corpo = montar_corpo_html(qtd_arquivos, data_ref, caminho_base)

    msg.attach(MIMEText(corpo, "html", "utf-8"))

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
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao ler arquivo 7z para anexo: {e}")
        _limpar_arquivo(caminho_7z)
        return False

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=TIMEOUT_SMTP) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatarios, msg.as_string())
        print(f"[EMAIL] Email enviado com sucesso para {', '.join(destinatarios)}")
        _limpar_arquivo(caminho_7z)
        return True
    except smtplib.SMTPAuthenticationError:
        _log_erro_email(arquivo_log, "[EMAIL] Erro de autenticação. Verifique o email e a senha de app no config.ini")
        _limpar_arquivo(caminho_7z)
        return False
    except Exception as e:
        _log_erro_email(arquivo_log, f"[EMAIL] Erro ao enviar email: {e}")
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
