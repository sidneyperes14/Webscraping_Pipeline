import sys
import traceback
from pathlib import Path

import extrator_dados
import extrator_dados3
import tratamento_dados
import tratamento_dados2
import tratamento_dados3


DOWNLOAD_DIR = Path.home() / "Downloads"


def _executar_etapa(numero: int, total: int, descricao: str, func, *args, **kwargs):
    print(f"{numero}/{total} {descricao}")
    resultado = func(*args, **kwargs)
    print(f"Concluído: {descricao}")
    return resultado


def _registrar_arquivos_para_limpeza(resultado, arquivos_para_limpar: list[Path]):
    if not resultado:
        return

    if isinstance(resultado, tuple):
        for item in resultado:
            if item is not None:
                arquivos_para_limpar.append(Path(item))
        return

    arquivos_para_limpar.append(Path(resultado))


def _limpar_arquivos_especificos(arquivos: list[Path]):
    if not arquivos:
        print("Nenhum arquivo específico foi registrado para exclusão.")
        return

    # remove duplicados preservando ordem
    vistos = set()
    arquivos_unicos = []
    for arquivo in arquivos:
        chave = str(arquivo.resolve()) if arquivo.exists() else str(arquivo)
        if chave not in vistos:
            vistos.add(chave)
            arquivos_unicos.append(arquivo)

    print("\nLimpando arquivos gerados nesta execução...")
    for arquivo in arquivos_unicos:
        try:
            if arquivo.exists() and arquivo.is_file():
                arquivo.unlink()
                print(f"Excluído: {arquivo.name}")
            else:
                print(f"Arquivo não encontrado para exclusão: {arquivo}")
        except Exception as e:
            print(f"Não foi possível excluir {arquivo.name}: {e}")


def main():
    print("Iniciando pipeline HomeZy...")
    print(f"Pasta monitorada: {DOWNLOAD_DIR}")

    arquivos_para_limpar: list[Path] = []

    etapas = [
        ("Extração Olist (relatórios 1, 2 e empresa 2) - extrator_dados.py", extrator_dados.run, {"headless": True}),
        ("Extração Uoou (extrator_dados3.py)", extrator_dados3.run, {"headless": True}),
        ("Tratamento Olist 1 (tratamento_dados.py)", tratamento_dados.run, {}),
        ("Tratamento Olist 2 (tratamento_dados2.py)", tratamento_dados2.run, {}),
        ("Tratamento Uoou (tratamento_dados3.py)", tratamento_dados3.run, {}),
    ]

    total = len(etapas)

    try:
        for i, (descricao, func, kwargs) in enumerate(etapas, start=1):
            resultado = _executar_etapa(i, total, descricao, func, **kwargs)

            if func is extrator_dados.run or func is extrator_dados3.run:
                _registrar_arquivos_para_limpeza(resultado, arquivos_para_limpar)

        print("Pipeline finalizado com sucesso!")
        _limpar_arquivos_especificos(arquivos_para_limpar)
        return 0

    except Exception as e:
        print(f"Pipeline falhou na etapa atual: {str(e)}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())