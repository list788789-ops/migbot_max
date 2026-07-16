"""Одноразовая диагностика: почему журнал инструктажей не конвертируется в PDF.
Читает реальные записи (read-only), генерит XLSX как прод, пробует soffice.
Ничего в БД не меняет. Удаляется после диагностики.
"""
import os, subprocess, tempfile, traceback
from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
import openpyxl

load_dotenv()
DATABASE_URL = os.environ.get("DATABASE_URL", "sqlite:///./migbot.db")
ORG_NAME = os.environ.get("COMPANY_NAME", "ИП Буц Сергей Юрьевич")
engine = create_engine(DATABASE_URL)

import production as prod
from models import InstructionType


def _soffice(xlsx_path, out_dir):
    soffice = "/usr/bin/soffice"
    profile = tempfile.mkdtemp(prefix="lo_diag_")
    proc = subprocess.run(
        [soffice, "--headless", "--nologo", "--nofirststartwizard",
         f"-env:UserInstallation=file://{profile}",
         "--convert-to", "pdf", "--outdir", out_dir, xlsx_path],
        capture_output=True, text=True, timeout=90,
    )
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf")
    ok = os.path.exists(pdf)
    tail = (proc.stderr or proc.stdout or "").strip()[-400:]
    return ok, proc.returncode, tail


def main():
    itype = InstructionType.INTRODUCTORY
    tmp = tempfile.mkdtemp(prefix="jdiag_")
    with Session(engine) as s:
        entries = prod.get_journaled_instructions(s, itype)
        order_ref = prod.get_latest_order_ref(s)
        started_at = prod.get_journal_started_at(s, itype)
        print("ENTRIES", len(entries), "ORDER_REF", repr(order_ref)[:40], "STARTED", started_at)
        xlsx = prod.generate_instruction_journal_xlsx(
            entries, itype, org_name=ORG_NAME, order_ref=order_ref,
            started_at=started_at, output_dir=tmp,
        )
    print("XLSX_PATH", xlsx, "SIZE", os.path.getsize(xlsx))

    # 1) открывается ли файл самим openpyxl (проверка целостности zip/xml)
    try:
        wb = openpyxl.load_workbook(xlsx)
        ws = wb.active
        print("OPENPYXL_RELOAD ok, sheet", ws.title, "row_breaks", len(ws.row_breaks.brk))
    except Exception as e:
        print("OPENPYXL_RELOAD FAIL", repr(e))
        wb = None

    # 2) как есть -> soffice
    ok, rc, tail = _soffice(xlsx, tmp)
    print("ASIS soffice ok", ok, "rc", rc, "tail", tail)

    # 3) без разрывов страниц -> soffice
    if wb is not None:
        try:
            ws.row_breaks.brk = []
            ws.col_breaks.brk = []
            nb_path = os.path.join(tmp, "journal_nobreaks.xlsx")
            wb.save(nb_path)
            ok2, rc2, tail2 = _soffice(nb_path, tmp)
            print("NOBREAKS soffice ok", ok2, "rc", rc2, "tail", tail2)
        except Exception:
            print("NOBREAKS step failed")
            traceback.print_exc()

    # 4) минимальный чистый xlsx -> soffice (контроль: работает ли soffice вообще с xlsx)
    try:
        wb2 = openpyxl.Workbook()
        wb2.active["A1"] = "hello"
        mp = os.path.join(tmp, "minimal.xlsx")
        wb2.save(mp)
        ok3, rc3, tail3 = _soffice(mp, tmp)
        print("MINIMAL soffice ok", ok3, "rc", rc3, "tail", tail3)
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
