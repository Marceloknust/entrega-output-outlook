from __future__ import annotations

import calendar
import logging
import shutil
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog

import pandas as pd
import win32com.client  # type: ignore[import]


BASE_DIR = Path(__file__).resolve().parents[1]
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OLD_DIR = OUTPUT_DIR / "old"

DE_PARA_FILE = INPUT_DIR / "DE_PARA.xlsx"
DE_PARA_SHEET = "DE_PARA"


def setup_logging() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    OLD_DIR.mkdir(parents=True, exist_ok=True)
    log_path = OUTPUT_DIR / f"script2_aplicar_ajustes_{datetime.now():%Y%m%d_%H%M%S}.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    return log_path


def normalize_name(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.strip().lower()
    text = text.replace("á", "a").replace("à", "a").replace("â", "a").replace("ã", "a")
    text = text.replace("é", "e").replace("ê", "e")
    text = text.replace("í", "i")
    text = text.replace("ó", "o").replace("ô", "o").replace("õ", "o")
    text = text.replace("ú", "u")
    text = text.replace("ç", "c")
    return " ".join(text.split())


def detect_column(df: pd.DataFrame, candidates: list[str]) -> str:
    normalized_map = {normalize_name(c): str(c) for c in df.columns}
    for name in candidates:
        key = normalize_name(name)
        if key in normalized_map:
            return normalized_map[key]
    raise ValueError(f"Coluna nao encontrada. Candidatas: {candidates}")


def normalize_integral(value: object) -> str:
    text = normalize_name(value)
    if text == "sim":
        return "sim"
    if text == "nao":
        return "nao"
    return ""


def resolve_de_para_sheet(path: Path) -> str:
    wb = pd.ExcelFile(path)
    for sheet in wb.sheet_names:
        normalized = normalize_name(sheet)
        if "de" in normalized and "para" in normalized:
            return sheet
    if DE_PARA_SHEET in wb.sheet_names:
        return DE_PARA_SHEET
    return wb.sheet_names[0]


def ask_validacao_file() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    default_name = "cronograma_outlook_validacao_"
    path_str = simpledialog.askstring(
        "Arquivo validado",
        "Informe o caminho completo do arquivo validado (.xlsx):",
        initialvalue=str(OUTPUT_DIR / f"{default_name}AAAAMM.xlsx"),
        parent=root,
    )
    root.destroy()

    if not path_str:
        return None

    path = Path(path_str.strip().strip('"'))
    if not path.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {path}")
    return path


def parse_yyyymm_from_filename(path: Path) -> str:
    stem = path.stem
    digits = "".join(ch for ch in stem if ch.isdigit())
    if len(digits) >= 6:
        return digits[-6:]
    now = datetime.now()
    return f"{now.year}{now.month:02d}"


def backup_de_para() -> Path:
    if not DE_PARA_FILE.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {DE_PARA_FILE}")

    OLD_DIR.mkdir(parents=True, exist_ok=True)
    backup_name = f"DE_PARA_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    backup_path = OLD_DIR / backup_name
    shutil.copy2(DE_PARA_FILE, backup_path)
    return backup_path


def detect_changes(df: pd.DataFrame) -> pd.DataFrame:
    for col in [
        "Atendimento",
        "Contrato",
        "Modalidade",
        "Atendimento_ajustado",
        "Contrato_ajustado",
        "Modalidade_ajustado",
        "Bacia Ajustada",
        "Bloco Ajustado",
        "Prefixo SAP",
    ]:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatoria ausente no cronograma_base: {col}")

    changed_mask = (
        (df["Atendimento"].fillna("").astype(str) != df["Atendimento_ajustado"].fillna("").astype(str))
        | (df["Contrato"].fillna("").astype(str) != df["Contrato_ajustado"].fillna("").astype(str))
        | (df["Modalidade"].fillna("").astype(str) != df["Modalidade_ajustado"].fillna("").astype(str))
    )
    return df.loc[changed_mask].copy()


def update_de_para(df_de_para: pd.DataFrame, df_changes: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    col_bacia = detect_column(df_de_para, ["Bacia Ajustado", "Bacia Ajustada"])
    col_bloco = detect_column(df_de_para, ["Bloco Ajustado"])
    col_unidade = detect_column(df_de_para, ["Unidade Marítima", "Unidade Maritima"])
    col_integral = detect_column(df_de_para, ["Integral?", "Integral"])
    col_atendimento = detect_column(df_de_para, ["Atendimento"])
    col_contrato = detect_column(df_de_para, ["Contrato"])
    col_modalidade = detect_column(df_de_para, ["Modalidade"])

    de_para = df_de_para.copy()
    de_para["_row_id"] = de_para.index
    de_para["_integral"] = de_para[col_integral].map(normalize_integral)
    de_para["_k_bacia"] = de_para[col_bacia].map(normalize_name)
    de_para["_k_bloco"] = de_para[col_bloco].map(normalize_name)
    de_para["_k_unid"] = de_para[col_unidade].map(normalize_name)

    invalid_integral = de_para[~de_para["_integral"].isin(["sim", "nao"])]
    if not invalid_integral.empty:
        raise ValueError(
            "DE_PARA possui linhas com Integral? inválido/vazio. "
            "Corrija antes de aplicar os ajustes."
        )

    sim_dup = (
        de_para[de_para["_integral"] == "sim"]
        .groupby(["_k_bacia", "_k_bloco"], dropna=False)
        .size()
        .reset_index(name="qtd")
    )
    sim_dup = sim_dup[sim_dup["qtd"] > 1]

    nao_dup = (
        de_para[de_para["_integral"] == "nao"]
        .groupby(["_k_bloco", "_k_unid"], dropna=False)
        .size()
        .reset_index(name="qtd")
    )
    nao_dup = nao_dup[nao_dup["qtd"] > 1]

    if not sim_dup.empty or not nao_dup.empty:
        details: list[str] = []
        for _, row in sim_dup.head(10).iterrows():
            details.append(f"SIM(bacia+bloco)=({row['_k_bacia']},{row['_k_bloco']}) qtd={row['qtd']}")
        for _, row in nao_dup.head(10).iterrows():
            details.append(f"NAO(bloco+unidade)=({row['_k_bloco']},{row['_k_unid']}) qtd={row['qtd']}")
        logging.warning(
            "DE_PARA possui chaves duplicadas para a regra de Integral?. "
            "Mantendo sempre a primeira ocorrência para atualização. %s",
            " | ".join(details),
        )

    de_para_lookup_sim = de_para[de_para["_integral"] == "sim"].drop_duplicates(
        subset=["_k_bacia", "_k_bloco"],
        keep="first",
    )
    de_para_lookup_nao = de_para[de_para["_integral"] == "nao"].drop_duplicates(
        subset=["_k_bloco", "_k_unid"],
        keep="first",
    )

    updated = 0
    not_found = 0

    grouped = (
        df_changes.assign(
            _k_bacia=df_changes["Bacia Ajustada"].map(normalize_name),
            _k_bloco=df_changes["Bloco Ajustado"].map(normalize_name),
            _k_unid=df_changes["Prefixo SAP"].map(normalize_name),
        )
        .drop_duplicates(subset=["_k_bacia", "_k_bloco", "_k_unid"], keep="last")
    )

    for _, row in grouped.iterrows():
        key_bacia = row["_k_bacia"]
        key_bloco = row["_k_bloco"]
        key_unid = row["_k_unid"]
        mask_nao = (
            (de_para_lookup_nao["_k_bloco"] == key_bloco)
            & (de_para_lookup_nao["_k_unid"] == key_unid)
        )

        # Regra: se existir caso de Integral=Não, nao considerar chave por Bacia+Bloco.
        if mask_nao.any():
            row_ids = de_para_lookup_nao.loc[mask_nao, "_row_id"]
        else:
            mask_sim = (
                (de_para_lookup_sim["_k_bacia"] == key_bacia)
                & (de_para_lookup_sim["_k_bloco"] == key_bloco)
            )
            row_ids = de_para_lookup_sim.loc[mask_sim, "_row_id"]

        if row_ids.empty:
            not_found += 1
            continue

        de_para.loc[row_ids, col_atendimento] = row["Atendimento_ajustado"]
        de_para.loc[row_ids, col_contrato] = row["Contrato_ajustado"]
        de_para.loc[row_ids, col_modalidade] = row["Modalidade_ajustado"]
        updated += int(len(row_ids))

    de_para = de_para.drop(columns=["_row_id", "_integral", "_k_bacia", "_k_bloco", "_k_unid"])
    return de_para, updated, not_found


def build_final_cronograma(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    for col in ["Atendimento", "Contrato", "Modalidade"]:
        output[col] = output[f"{col}_ajustado"]

    drop_cols = [
        "InicioNoMes_ajustado",
        "FimNoMes_ajustado",
        "Atendimento_ajustado",
        "Contrato_ajustado",
        "Modalidade_ajustado",
        "Validado",
    ]
    existing_drop_cols = [c for c in drop_cols if c in output.columns]
    output = output.drop(columns=existing_drop_cols)
    return output


def build_td(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    td = pd.pivot_table(
        df,
        index=["Bloco Ajustado", "Prefixo SAP"],
        columns=["MesReferencia"],
        values="DiasOperadosMes",
        aggfunc="sum",
        fill_value=0,
    )
    td = td.reset_index()
    td.columns = [str(c) for c in td.columns]
    return td


def prompt_de_para_overwrite() -> bool:
    root = tk.Tk()
    root.withdraw()
    confirm = messagebox.askyesno(
        "Confirmar atualizacao",
        "Deseja atualizar o DE_PARA.xlsx com as alteracoes encontradas?",
        parent=root,
    )
    root.destroy()
    return confirm


def send_email_final(attachment: Path, yyyymm: str) -> None:
    year = int(yyyymm[:4])
    month = int(yyyymm[4:])

    meses_pt = [
        "Janeiro",
        "Fevereiro",
        "Marco",
        "Abril",
        "Maio",
        "Junho",
        "Julho",
        "Agosto",
        "Setembro",
        "Outubro",
        "Novembro",
        "Dezembro",
    ]

    default_subject = f"Cronograma Final Consolidado - {meses_pt[month - 1]}/{year}"
    default_to = "felipedan@transpetro.com.br"
    default_body = (
        "Felipe,\n\n"
        f"Segue arquivo \"{attachment.name}\" com as validacoes do time Comercial consolidadas.\n\n"
        "O DE_PARA foi atualizado com as alteracoes realizadas.\n\n"
        "Qualquer duvida, estou a disposicao.\n\n"
        "Atenciosamente,\n"
        "Guilherme"
    )

    root = tk.Tk()
    root.withdraw()

    to_value = simpledialog.askstring(
        "Destinatario",
        "E-mail Para:",
        initialvalue=default_to,
        parent=root,
    )
    if to_value is None:
        root.destroy()
        logging.info("Envio cancelado pelo usuario.")
        return

    subject_value = simpledialog.askstring(
        "Assunto",
        "Assunto do e-mail:",
        initialvalue=default_subject,
        parent=root,
    )
    if subject_value is None:
        root.destroy()
        logging.info("Envio cancelado pelo usuario.")
        return

    body_value = simpledialog.askstring(
        "Corpo",
        "Corpo do e-mail:",
        initialvalue=default_body,
        parent=root,
    )
    if body_value is None:
        root.destroy()
        logging.info("Envio cancelado pelo usuario.")
        return

    should_send = messagebox.askyesno("Confirmar envio", "Deseja enviar o e-mail final agora?", parent=root)
    root.destroy()
    if not should_send:
        logging.info("Envio final cancelado pelo usuario.")
        return

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to_value
    mail.Subject = subject_value
    mail.Body = body_value
    mail.Attachments.Add(str(attachment.resolve()))
    mail.Send()
    logging.info("E-mail final enviado para: %s", to_value)


def main() -> None:
    log_path = setup_logging()
    logger = logging.getLogger("script2")
    logger.info("Inicio do script2_aplicar_ajustes")
    logger.info("Log: %s", log_path)

    try:
        validacao_path = ask_validacao_file()
        if validacao_path is None:
            logger.info("Execucao cancelada pelo usuario.")
            return

        yyyymm = parse_yyyymm_from_filename(validacao_path)

        logger.info("Arquivo validado: %s", validacao_path)
        df_cronograma = pd.read_excel(validacao_path, sheet_name="cronograma_base")

        # Base Dados/TD podem nao existir em arquivos alterados manualmente.
        try:
            df_base_dados = pd.read_excel(validacao_path, sheet_name="Base Dados")
        except Exception:
            df_base_dados = pd.DataFrame()

        df_changes = detect_changes(df_cronograma)
        logger.info("Alteracoes encontradas: %d", len(df_changes))

        if not prompt_de_para_overwrite():
            logger.info("Atualizacao de DE_PARA cancelada pelo usuario.")
            return

        backup_path = backup_de_para()
        logger.info("Backup DE_PARA criado: %s", backup_path)

        de_para_sheet = resolve_de_para_sheet(DE_PARA_FILE)
        logger.info("Aba DE_PARA selecionada: %s", de_para_sheet)
        df_de_para = pd.read_excel(DE_PARA_FILE, sheet_name=de_para_sheet)
        df_de_para_updated, updated_rows, not_found = update_de_para(df_de_para, df_changes)

        with pd.ExcelWriter(DE_PARA_FILE, engine="openpyxl") as writer:
            df_de_para_updated.to_excel(writer, sheet_name=de_para_sheet, index=False)

        logger.info("Linhas atualizadas no DE_PARA: %d", updated_rows)
        if not_found:
            logger.warning("Chaves nao encontradas no DE_PARA: %d", not_found)

        df_final = build_final_cronograma(df_cronograma)
        df_td = build_td(df_final)

        output_final = OUTPUT_DIR / f"cronograma_outlook_final_{yyyymm}.xlsx"
        with pd.ExcelWriter(output_final, engine="openpyxl", datetime_format="DD/MM/YYYY") as writer:
            df_final.to_excel(writer, sheet_name="cronograma_final", index=False)
            df_td.to_excel(writer, sheet_name="TD", index=False)
            if not df_base_dados.empty:
                df_base_dados.to_excel(writer, sheet_name="Base Dados", index=False)

        logger.info("Arquivo final gerado: %s", output_final)

        send_email_final(output_final, yyyymm)
        logger.info("Script2 finalizado com sucesso.")
    except Exception as exc:
        logger.exception("Falha na execucao: %s", exc)
        raise


if __name__ == "__main__":
    main()
