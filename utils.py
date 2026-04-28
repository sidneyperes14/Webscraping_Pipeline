"""
Utilitários compartilhados pelos scripts de tratamento.

Centraliza helpers de:
- espera por liberação de arquivo (lock do OneDrive/Excel)
- re-salvamento do arquivo via Excel COM (normaliza XLSX para o Power BI)
- encerramento de processos EXCEL.EXE zumbis
"""
from __future__ import annotations

import os
import time
import subprocess
from pathlib import Path


def matar_excel_zumbi() -> None:
    """
    Encerra processos EXCEL.EXE remanescentes.

    Útil antes de operações via win32com para evitar conflitos
    com execuções anteriores que crasharam.
    """
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE"],
            capture_output=True,
            timeout=10,
        )
    except Exception:
        pass


def esperar_arquivo_disponivel(
    path: Path,
    timeout_sec: int = 120,
    intervalo: float = 2.0,
) -> None:
    """
    Aguarda até que o arquivo possa ser aberto para escrita.

    Cobre casos de:
    - OneDrive/SharePoint sincronizando o arquivo
    - Excel aberto em outro processo
    - Arquivo bloqueado temporariamente por antivirus
    """
    if not path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    fim = time.time() + timeout_sec
    ultima_falha = None

    while time.time() < fim:
        try:
            # Tenta abrir em modo append binário — requer lock de escrita
            with open(path, "r+b"):
                return
        except (PermissionError, OSError) as e:
            ultima_falha = e
            time.sleep(intervalo)

    raise PermissionError(
        f"Arquivo continuou bloqueado após {timeout_sec}s: {path}\n"
        f"Último erro: {ultima_falha}\n"
        "Verifique se o arquivo não está aberto no Excel ou sincronizando no OneDrive."
    )


def resalvar_via_excel(xlsx_path: Path, max_attempts: int = 5) -> None:
    """
    Abre o arquivo no Excel (COM) e salva novamente.

    Força o Excel a regravar o arquivo com fórmulas recalculadas,
    valores cacheados atualizados e metadados completos — corrige
    inconsistências deixadas pelo openpyxl que causam erro no Power BI.
    """
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "Para normalizar o XLSX via Excel, instale o pywin32:\n"
            "  pip install pywin32"
        ) from e

    last_err = None

    for attempt in range(1, max_attempts + 1):
        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(str(xlsx_path))
            excel.CalculateFullRebuild()
            wb.Save()
            wb.Close(SaveChanges=False)
            wb = None

            excel.Quit()
            excel = None
            return
        except Exception as e:
            last_err = e
            # se o arquivo está travado, espera um pouco e tenta de novo
            time.sleep(3)
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    raise RuntimeError(
        f"Falha ao re-salvar {xlsx_path} via Excel após {max_attempts} tentativas: {last_err}"
    )
