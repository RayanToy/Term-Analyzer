"""
Финальный шаг: PDF размеченных учебников через Word.

Числа упоминаний у терминов стоят в плавающих надписях, которые не
занимают места в строке, поэтому полей и их ошибок больше нет — документ
просто экспортируется в PDF. После экспорта проверяется, что в PDF нет
«Ошибка!» (на случай полей, оставшихся в исходнике).

Word запускается отдельным процессом через PowerShell (COM); открытые
пользователем окна Word не затрагиваются.

Запуск: python word_finalize.py <docx> [<docx> ...]
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pypdfium2 as pdfium

ERROR_MARK = "Ошибка!"

PS_EXPORT = r"""
$ErrorActionPreference = 'Stop'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$word = New-Object -ComObject Word.Application
Start-Sleep -Milliseconds 800
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Where-Object { $before -notcontains $_ })
if ($mine.Count -eq 0) { throw 'Не удалось запустить отдельный процесс Word' }
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    foreach ($d in $args) {
        $pdf = [System.IO.Path]::ChangeExtension($d, '.pdf')
        $doc = $word.Documents.Open($d, $false, $true)
        $doc.ExportAsFixedFormat($pdf, 17)
        $doc.Close(0)
        Write-Output "PDF: $pdf"
    }
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Start-Sleep -Seconds 2
    foreach ($id in $mine) { Get-Process -Id $id -ErrorAction SilentlyContinue | Stop-Process -Force }
}
"""


def error_pages(pdf_path: Path) -> list[int]:
    doc = pdfium.PdfDocument(str(pdf_path))
    pages = [i + 1 for i in range(len(doc)) if ERROR_MARK in doc[i].get_textpage().get_text_bounded()]
    doc.close()
    return pages


def export_pdf(paths: list[Path]) -> None:
    script = "& {" + PS_EXPORT + "} " + " ".join("'" + str(p).replace("'", "''") + "'" for p in paths)
    res = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                         capture_output=True, text=True, encoding="cp866", errors="ignore")
    print(res.stdout.strip())
    if res.returncode:
        raise SystemExit(f"[Word] экспорт не удался: {res.stderr.strip()}")


def main(paths: list[str]) -> None:
    # «~$…docx» — служебные файлы блокировки открытых в Word документов
    docs = [Path(p).resolve() for p in paths if not Path(p).name.startswith("~$")]
    export_pdf(docs)
    for d in docs:
        pdf = d.with_suffix(".pdf")
        doc = pdfium.PdfDocument(str(pdf))
        n = len(doc)
        doc.close()
        bad = error_pages(pdf)
        print(f"[Word] {pdf.name}: {n} стр." + (f"; «{ERROR_MARK}» на стр. {bad}" if bad else ", ошибок нет"))


if __name__ == "__main__":
    main(sys.argv[1:])
