import sys
import traceback

import extrator_dados
import extrator_dados3
import tratamento_dados
import tratamento_dados2
import tratamento_dados3


def _executar_etapa(numero: int, total: int, descricao: str, func, *args, **kwargs):
    print(f"{numero}/{total} {descricao}")
    func(*args, **kwargs)
    print(f"Concluído: {descricao}")


def main():
    print("Iniciando pipeline HomeZy...")

    etapas = [
        ("Extração Olist (relatórios 1 e 2) - extrator_dados.py", extrator_dados.run, {"headless": True}),
        ("Extração Uoou (extrator_dados3.py)", extrator_dados3.run, {"headless": True}),
        ("Tratamento Olist 1 (tratamento_dados.py)", tratamento_dados.run, {}),
        ("Tratamento Olist 2 (tratamento_dados2.py)", tratamento_dados2.run, {}),
        ("Tratamento Uoou (tratamento_dados3.py)", tratamento_dados3.run, {}),
    ]

    total = len(etapas)

    try:
        for i, (descricao, func, kwargs) in enumerate(etapas, start=1):
            _executar_etapa(i, total, descricao, func, **kwargs)

        print("Pipeline finalizado com sucesso!")
        return 0

    except Exception as e:
        print(f"Pipeline falhou na etapa atual: {str(e)}")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())