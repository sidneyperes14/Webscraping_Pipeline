import os
import time
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException

from webdriver_manager.chrome import ChromeDriverManager

import config  # usa DOWNLOADS_DIR


def xls_to_xlsx_via_excel(xls_path: Path, xlsx_path: Path, max_attempts: int = 5) -> Path:
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "Para converter XLS -> XLSX automaticamente, instale o pywin32:\n"
            "  pip install pywin32\n"
            "Depois reinicie o VS Code/terminal."
        ) from e

    last_err = None

    for _attempt in range(1, max_attempts + 1):
        excel = None
        wb = None
        try:
            pythoncom.CoInitialize()

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            xlsx_path.parent.mkdir(parents=True, exist_ok=True)

            if xlsx_path.exists():
                try:
                    xlsx_path.unlink()
                except PermissionError:
                    raise RuntimeError(
                        f"Não consegui sobrescrever o XLSX: {xlsx_path}\n"
                        "Feche o Excel/preview que esteja usando esse arquivo."
                    )

            wb = excel.Workbooks.Open(str(xls_path))
            wb.SaveAs(str(xlsx_path), FileFormat=51)
            wb.Close(SaveChanges=False)
            wb = None

            excel.Quit()
            excel = None
            return xlsx_path

        except Exception as e:
            last_err = e
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
            time.sleep(2)

        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    raise RuntimeError(
        f"Falha ao converter {xls_path} para XLSX após {max_attempts} tentativas: {last_err}"
    )


def run(headless: bool = True) -> Path:
    load_dotenv()
    url = os.getenv("UOOU_URL")
    username = os.getenv("UOOU_USERNAME")
    password = os.getenv("UOOU_PASSWORD")

    if not url or not username or not password:
        raise RuntimeError(
            "Faltou configurar UOOU_URL, UOOU_USERNAME ou UOOU_PASSWORD no .env"
        )

    download_dir = Path(config.DOWNLOADS_DIR).resolve()
    download_dir.mkdir(parents=True, exist_ok=True)

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")

    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-blink-features=AutomationControlled")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )

    wait = WebDriverWait(driver, 60)

    def esperar_overlay_sumir(timeout: int = 30) -> None:
        overlays = [
            (By.CSS_SELECTOR, ".displayWaitBlock.waitDataConfigWidget"),
            (By.CSS_SELECTOR, ".displayWaitBlock"),
        ]
        for locator in overlays:
            try:
                WebDriverWait(driver, timeout).until(
                    EC.invisibility_of_element_located(locator)
                )
            except TimeoutException:
                pass

    def clicar(xpath: str, timeout: int = 60) -> None:
        esperar_overlay_sumir()
        elem = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        time.sleep(0.5)

        try:
            elem.click()
        except ElementClickInterceptedException:
            esperar_overlay_sumir()
            time.sleep(1)
            elem = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            driver.execute_script("arguments[0].click();", elem)

    def preencher(xpath: str, valor: str, timeout: int = 60) -> None:
        elem = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH, xpath))
        )
        elem.clear()
        elem.send_keys(valor)

    def trocar_para_nova_aba(qtd_abas_antes: int, timeout: int = 10) -> None:
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.window_handles) > qtd_abas_antes
        )
        driver.switch_to.window(driver.window_handles[-1])

    def clicar_dia_flatpickr(xpath_input: str, data_alvo: datetime, timeout: int = 60) -> None:
        meses_pt = {
            1: "Janeiro",
            2: "Fevereiro",
            3: "Março",
            4: "Abril",
            5: "Maio",
            6: "Junho",
            7: "Julho",
            8: "Agosto",
            9: "Setembro",
            10: "Outubro",
            11: "Novembro",
            12: "Dezembro",
        }
        meses_pt_inv = {v: k for k, v in meses_pt.items()}

        aria_label = f"{meses_pt[data_alvo.month]} {data_alvo.day}, {data_alvo.year}"
        seletor_dia = f'span.flatpickr-day[aria-label="{aria_label}"]'

        clicar(xpath_input, timeout=timeout)
        time.sleep(0.5)

        for _ in range(24):
            dias = driver.find_elements(By.CSS_SELECTOR, seletor_dia)
            if dias:
                dia_elem = dias[0]
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dia_elem)
                time.sleep(0.2)
                try:
                    dia_elem.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", dia_elem)
                time.sleep(0.5)
                return

            mes_visivel = driver.find_element(
                By.CSS_SELECTOR, ".flatpickr-current-month .cur-month"
            ).text.strip()

            ano_visivel = int(
                driver.find_element(
                    By.CSS_SELECTOR, ".flatpickr-current-month .numInput.cur-year"
                ).get_attribute("value").strip()
            )

            mes_visivel_num = meses_pt_inv[mes_visivel]

            if (ano_visivel, mes_visivel_num) > (data_alvo.year, data_alvo.month):
                driver.find_element(By.CSS_SELECTOR, ".flatpickr-prev-month").click()
            else:
                driver.find_element(By.CSS_SELECTOR, ".flatpickr-next-month").click()

            time.sleep(0.4)

        raise TimeoutException(f"Não consegui localizar a data no calendário: {aria_label}")

    def clicar_download_relatorio_mais_recente(timeout: int = 600) -> None:
        """
        Espera a lista de relatorios carregar e clica no botao de download
        do relatorio mais recente.

        Se o relatorio mais novo aparecer no final da tabela,
        troque usar_primeiro para False.
        """
        usar_primeiro = True

        xpath_botoes = '//table//tr[.//td[9]//a]//td[9]//a//span'

        fim = time.time() + timeout
        ultima_qtd = 0

        while time.time() < fim:
            try:
                esperar_overlay_sumir(timeout=5)
            except Exception:
                pass

            botoes = driver.find_elements(By.XPATH, xpath_botoes)
            botoes_visiveis = [b for b in botoes if b.is_displayed()]

            if botoes_visiveis:
                ultima_qtd = len(botoes_visiveis)
                alvo = botoes_visiveis[0] if usar_primeiro else botoes_visiveis[-1]

                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", alvo)
                time.sleep(0.5)

                try:
                    alvo.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", alvo)

                print(f"Botão de download localizado. Total de botões encontrados: {ultima_qtd}")
                return

            time.sleep(10)
            driver.refresh()

        raise TimeoutException(
            f"Não encontrei botão de download na grade de relatórios dentro de {timeout}s."
        )

    def esperar_download_apos_clique(timeout_sec: int = 300) -> Path:
        before_files = {
            p.name for p in download_dir.iterdir()
            if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".csv"}
        }

        clicar_download_relatorio_mais_recente(timeout=600)

        start = time.time()
        downloaded_file = None
        stable_count = 0
        last_size = None

        while time.time() - start < timeout_sec:
            time.sleep(1)

            if list(download_dir.glob("*.crdownload")):
                stable_count = 0
                last_size = None
                continue

            files = [
                p for p in download_dir.iterdir()
                if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".csv"}
            ]

            new_files = [p for p in files if p.name not in before_files]

            if not new_files:
                continue

            candidate = max(new_files, key=lambda p: p.stat().st_mtime)

            try:
                size = candidate.stat().st_size
            except FileNotFoundError:
                continue

            if size <= 0:
                continue

            if last_size is None or size != last_size:
                last_size = size
                stable_count = 0
            else:
                stable_count += 1

            if stable_count >= 5:
                downloaded_file = candidate
                break

        if not downloaded_file:
            raise TimeoutError(
                f"Download não finalizou em {timeout_sec}s na pasta {download_dir}"
            )

        time.sleep(2)
        print(f"Download concluído: {downloaded_file}")
        return downloaded_file

    def renomear_e_converter(downloaded_file: Path) -> Path:
        yyyymm = datetime.now().strftime("%Y%m")
        ext = downloaded_file.suffix.lower()

        if ext not in {".xls", ".xlsx"}:
            raise RuntimeError(
                f"O arquivo baixado veio com extensão inesperada: {downloaded_file.name}"
            )

        final_name = f"uoou_relatorio_{yyyymm}{ext}"
        final_path = download_dir / final_name

        if final_path.exists():
            try:
                final_path.unlink()
            except PermissionError:
                raise RuntimeError(
                    f"Não consegui substituir {final_path}.\n"
                    "Provavelmente ele está aberto no Excel."
                )

        downloaded_file.rename(final_path)
        print(f"Arquivo renomeado: {final_path}")

        if final_path.suffix.lower() == ".xls":
            xlsx_final_path = final_path.with_suffix(".xlsx")
            print(f"Convertendo para XLSX: {xlsx_final_path}")
            xls_to_xlsx_via_excel(final_path, xlsx_final_path, max_attempts=5)
            print(f"XLSX gerado: {xlsx_final_path}")
            return xlsx_final_path

        print(f"Arquivo já veio em XLSX: {final_path}")
        return final_path

    try:
        driver.get(url)

        # login
        preencher('//*[@id="_username"]', username)
        preencher('//*[@id="_password"]', password)
        clicar('//*[@id="send"]')

        # menu
        clicar('//*[@id="sidebar"]/div/div/ul/li[3]/ul/li[3]/a/span')

        # datas via calendário flatpickr
        hoje = datetime.now()
        primeiro_dia_ano = datetime(hoje.year, 1, 1)

        clicar_dia_flatpickr('//*[@id="filterForm"]/div/div[1]/div/input[2]', primeiro_dia_ano)
        clicar_dia_flatpickr('//*[@id="filterForm"]/div/div[2]/div/input[2]', hoje)

        # filtrar
        clicar('//*[@id="filterForm"]/div/div[24]/div/button')

        # exportar fretes
        clicar('//*[@id="content"]/div[2]/div/div[1]/button')

        abas_antes = len(driver.window_handles)
        clicar('//*[@id="content"]/div[2]/div/div[1]/ul/li[1]/a')
        trocar_para_nova_aba(abas_antes)

        # aguarda o relatório ficar pronto, mantendo a sessão ativa
        espera_total = 485
        intervalo = 30
        print(f"Aguardando {espera_total // 60} minutos e {espera_total % 60} segundos para geração do relatório...")
        inicio_espera = time.time()
        while time.time() - inicio_espera < espera_total:
            time.sleep(min(intervalo, espera_total - (time.time() - inicio_espera)))
            try:
                driver.current_url  # keep-alive: evita timeout da sessão
            except Exception:
                pass

        # localizar o botão do relatório mais recente e baixar
        downloaded_file = esperar_download_apos_clique(timeout_sec=300)

        arquivo_final = renomear_e_converter(downloaded_file)
        return arquivo_final

    finally:
        driver.quit()


if __name__ == "__main__":
    run(headless=True)