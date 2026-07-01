"""
output_outlook — Automação de Demanda Primária via Outlook + Excel
=================================================================
Fluxo:
  1. Lê arquivos .xlsx da pasta input/ (planilhas de demanda primária).
  2. Consolida os dados em um único DataFrame.
  3. Salva o consolidado em output/ como Excel.
  4. Envia o arquivo gerado por e-mail via Outlook (win32com).
  5. (Opcional) Busca e-mails com anexos Excel na caixa de entrada e
     salva os anexos automaticamente em input/.

Pré-requisitos:
  - Windows com Outlook instalado e conta configurada.
  - pip install pandas openpyxl pywin32
"""

from __future__ import annotations

import logging
from datetime import date
from pathlib import Path

import pandas as pd
import win32com.client  # type: ignore[import]

BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# --- Configurações de e-mail ---
DESTINATARIOS = ["destinatario@empresa.com"]   # ajuste conforme necessário
ASSUNTO_ENVIO = f"Demanda Primária Consolidada — {date.today():%d/%m/%Y}"
CORPO_ENVIO = (
    "Olá,\n\n"
    "Segue em anexo o consolidado de Demanda Primária gerado automaticamente.\n\n"
    "Atenciosamente,\nAutomação output_outlook"
)

# Filtro para buscar e-mails recebidos (deixe vazio para não buscar)
ASSUNTO_RECEBIMENTO = "Demanda Primária"

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 1. Leitura e consolidação dos Excel de entrada
# ---------------------------------------------------------------------------

def consolidar_excel(input_dir: Path) -> pd.DataFrame | None:
    arquivos = list(input_dir.glob("*.xlsx"))
    if not arquivos:
        log.warning("Nenhum arquivo .xlsx encontrado em %s", input_dir)
        return None

    partes: list[pd.DataFrame] = []
    for arq in arquivos:
        log.info("Lendo: %s", arq.name)
        try:
            df = pd.read_excel(arq, engine="openpyxl")
            df["_origem"] = arq.name
            partes.append(df)
        except Exception as exc:
            log.error("Erro ao ler %s: %s", arq.name, exc)

    if not partes:
        return None

    consolidado = pd.concat(partes, ignore_index=True)
    log.info("Consolidado: %d linhas de %d arquivo(s).", len(consolidado), len(partes))
    return consolidado


# ---------------------------------------------------------------------------
# 2. Exportar para Excel na pasta output/
# ---------------------------------------------------------------------------

def salvar_excel(df: pd.DataFrame, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    nome = f"demanda_primaria_consolidada_{date.today():%Y%m%d}.xlsx"
    caminho = output_dir / nome
    df.to_excel(caminho, index=False, engine="openpyxl")
    log.info("Arquivo salvo: %s", caminho)
    return caminho


# ---------------------------------------------------------------------------
# 3. Enviar e-mail com anexo via Outlook
# ---------------------------------------------------------------------------

def enviar_email(caminho_anexo: Path) -> None:
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.Subject = ASSUNTO_ENVIO
        mail.Body = CORPO_ENVIO
        for dest in DESTINATARIOS:
            mail.Recipients.Add(dest)
        mail.Attachments.Add(str(caminho_anexo.resolve()))
        mail.Send()
        log.info("E-mail enviado para: %s", ", ".join(DESTINATARIOS))
    except Exception as exc:
        log.error("Falha ao enviar e-mail: %s", exc)


# ---------------------------------------------------------------------------
# 4. (Opcional) Baixar anexos Excel de e-mails recebidos → input/
# ---------------------------------------------------------------------------

def baixar_anexos_recebidos(input_dir: Path, assunto_filtro: str) -> int:
    if not assunto_filtro:
        return 0

    input_dir.mkdir(parents=True, exist_ok=True)
    baixados = 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        caixa = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        mensagens = caixa.Items
        mensagens.Sort("[ReceivedTime]", True)  # mais recentes primeiro

        for msg in mensagens:
            try:
                if assunto_filtro.lower() not in (msg.Subject or "").lower():
                    continue
                for anexo in msg.Attachments:
                    nome = Path(anexo.FileName)
                    if nome.suffix.lower() == ".xlsx":
                        destino = input_dir / nome.name
                        anexo.SaveAsFile(str(destino.resolve()))
                        log.info("Anexo salvo: %s", destino.name)
                        baixados += 1
            except Exception:
                continue
    except Exception as exc:
        log.error("Erro ao acessar caixa de entrada: %s", exc)

    return baixados


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------

def main() -> None:
    log.info("=== output_outlook iniciado ===")
    log.info("Input : %s", INPUT_DIR)
    log.info("Output: %s", OUTPUT_DIR)

    # Etapa 0 (opcional): buscar e-mails recebidos e salvar anexos em input/
    baixados = baixar_anexos_recebidos(INPUT_DIR, ASSUNTO_RECEBIMENTO)
    if baixados:
        log.info("%d anexo(s) baixado(s) da caixa de entrada.", baixados)

    # Etapa 1: consolidar Excel de input/
    df = consolidar_excel(INPUT_DIR)
    if df is None:
        log.warning("Nenhum dado para processar. Coloque arquivos .xlsx em input/ e rode novamente.")
        return

    # Etapa 2: salvar consolidado
    caminho = salvar_excel(df, OUTPUT_DIR)

    # Etapa 3: enviar por e-mail
    enviar_email(caminho)

    log.info("=== Concluído ===")


if __name__ == "__main__":
    main()
