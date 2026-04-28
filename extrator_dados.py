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
from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
)

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


def run(headless: bool = True) -> tuple[Path, Path, Path]:
    load_dotenv()

    url = os.getenv("OLIST_URL")
    username_1 = os.getenv("OLIST_USERNAME")
    password_1 = os.getenv("OLIST_PASSWORD")
    username_2 = os.getenv("OLIST2_USERNAME")
    password_2 = os.getenv("OLIST2_PASSWORD")

    if not url or not username_1 or not password_1:
        raise RuntimeError("Faltou configurar OLIST_URL, OLIST_USERNAME ou OLIST_PASSWORD no .env")

    if not username_2 or not password_2:
        raise RuntimeError("Faltou configurar OLIST2_USERNAME ou OLIST2_PASSWORD no .env")

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

    def confirmar_sessao_se_preciso() -> None:
        x_confirmar_sessao = '//*[@id="bs-modal-ui-popup"]/div/div/div/div[3]/button[1]'
        try:
            confirmar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, x_confirmar_sessao))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", confirmar)
            try:
                confirmar.click()
            except ElementClickInterceptedException:
                esperar_overlay_sumir()
                driver.execute_script("arguments[0].click();", confirmar)
            print("Sessão já ativa detectada: confirmei o acesso.")
        except TimeoutException:
            pass

    def fechar_modal_se_existir() -> None:
        candidatos = [
            '//*[@id="bs-modal"]/div/div/div/div[1]/button',
            '//*[@id="bs-modal-ui-popup"]/div/div/div/div[1]/button',
        ]
        for xp in candidatos:
            try:
                btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                )
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
            except TimeoutException:
                pass

    def existe_elemento(xpath: str, timeout: int = 5) -> bool:
        try:
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            return True
        except TimeoutException:
            return False

    def fazer_login(user: str, pwd: str, rotulo: str) -> None:
        print(f"Realizando login no Olist: {rotulo}")
        driver.get(url)

        preencher('//*[@id="username"]', user)
        preencher('//*[@id="password"]', pwd)

        btn_login = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, '//*[@id="kc-content-wrapper"]/react-login-wc/section/div/main/aside[1]/div/form/button')
            )
        )
        driver.execute_script('arguments[0].scrollIntoView({block:"center"});', btn_login)
        try:
            btn_login.click()
        except ElementClickInterceptedException:
            esperar_overlay_sumir()
            driver.execute_script("arguments[0].click();", btn_login)

        confirmar_sessao_se_preciso()
        esperar_overlay_sumir()
        fechar_modal_se_existir()

    def fazer_logout() -> None:
        print("Realizando logout da sessão atual...")
        fechar_modal_se_existir()
        esperar_overlay_sumir()

        x_menu_usuario = '//*[@id="main-menu"]/div[2]/div[1]/div[1]/div[2]/ul/li[5]/a/div/div[1]/div'
        clicar(x_menu_usuario, timeout=30)
        time.sleep(2)
        esperar_overlay_sumir()

        candidatos_logout = [
            '//*[@id="main-menu"]/div[2]/div[2]/nav[6]/div[3]/a',
            '//*[@id="main-menu"]/div[2]/div[2]//a[contains(@href, "logout")]',
            '//*[@id="main-menu"]/div[2]/div[2]//a[contains(., "Sair")]',
            '//*[@id="main-menu"]/div[2]/div[2]//a[contains(., "Logout")]',
            '//a[contains(@href, "logout")]',
            '//a[contains(., "Sair")]',
            '//a[contains(., "Logout")]',
        ]

        ultimo_erro = None

        for xp in candidatos_logout:
            try:
                elem = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.XPATH, xp))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
                time.sleep(0.5)
                esperar_overlay_sumir()

                try:
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, xp))
                    )
                    elem.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", elem)

                time.sleep(3)

                if existe_elemento('//*[@id="username"]', timeout=10):
                    print(f"Logout realizado com sucesso usando XPath: {xp}")
                    return

                print(f"Clique de logout executado usando XPath: {xp}")
                return

            except Exception as e:
                ultimo_erro = e

        raise RuntimeError(f"Não foi possível realizar o logout. Último erro: {ultimo_erro}")

    def snapshot_downloads() -> set[str]:
        return {
            p.name for p in download_dir.iterdir()
            if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".csv"}
        }

    def esperar_download(before_files: set[str], timeout_sec: int = 300) -> Path:
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
            raise TimeoutError(f"Download não finalizou em {timeout_sec}s na pasta {download_dir}")

        time.sleep(2)
        print(f"Download concluído: {downloaded_file}")
        return downloaded_file

    def renomear_e_converter(downloaded_file: Path, prefixo_final: str) -> Path:
        yyyymm = datetime.now().strftime("%Y%m")
        ext = downloaded_file.suffix.lower()

        if ext not in {".xls", ".xlsx"}:
            raise RuntimeError(
                f"O arquivo baixado veio com extensão inesperada: {downloaded_file.name}"
            )

        final_name = f"{prefixo_final}_{yyyymm}{ext}"
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

    def abrir_menu_principal() -> None:
        menu_xpath = '//*[@id="main-menu"]/div[2]/div[1]/div[1]/div[1]'
        try:
            clicar(menu_xpath, timeout=20)
            return
        except TimeoutException:
            print("Menu principal não ficou acessível. Recarregando a home do Olist...")
            driver.get(url)
            confirmar_sessao_se_preciso()
            esperar_overlay_sumir()
            fechar_modal_se_existir()
            clicar(menu_xpath, timeout=40)

    def baixar_relatorio_1() -> Path:
        print("Iniciando download do Relatório 1 - Empresa 1...")
        abrir_menu_principal()
        clicar('//*[@id="main-menu"]/div[2]/div[1]/div[1]/nav/ul/li[4]/a/span')
        clicar('//*[@id="main-menu"]/div[2]/div[2]/nav[4]/ul/li[7]/a/span')
        clicar('//*[@id="sit-E"]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[3]/ul/li[1]/a')
        clicar('//*[@id="opc-per-mes"]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[3]/ul/li[1]/div/div[5]/button[1]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[1]/div/div[2]/button/span[1]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[1]/div/div[2]/ul/li[15]/a')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[2]/div[2]/label')
        clicar('//*[@id="btnParamsExportarNotas"]')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[4]/div[3]/div/label')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[4]/div[4]/div[1]/label')

        before_files = snapshot_downloads()
        clicar('//*[@id="btnExportarNotas"]')

        downloaded_file = esperar_download(before_files, timeout_sec=600)
        return renomear_e_converter(downloaded_file, "olist_relatorio")

    def baixar_relatorio_1_empresa2() -> Path:
        print("Iniciando download do Relatório 1 - Empresa 2...")
        abrir_menu_principal()
        clicar('//*[@id="main-menu"]/div[2]/div[1]/div[1]/nav/ul/li[4]/a/span')
        clicar('//*[@id="main-menu"]/div[2]/div[2]/nav[4]/ul/li[7]/a/span')
        clicar('//*[@id="sit-E"]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[3]/ul/li[1]/a')
        clicar('//*[@id="opc-per-mes"]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[3]/ul/li[1]/div/div[5]/button[1]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[1]/div/div[2]/button/span[1]')
        clicar('//*[@id="page-wrapper"]/div[4]/div[1]/div[1]/div/div[2]/ul/li[13]/a')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[2]/div[2]/label')
        clicar('//*[@id="btnParamsExportarNotas"]')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[4]/div[3]/div/label')
        clicar('//*[@id="bs-modal"]/div/div/div/div[2]/div[4]/div[4]/div[1]/label')

        before_files = snapshot_downloads()
        clicar('//*[@id="btnExportarNotas"]')

        downloaded_file = esperar_download(before_files, timeout_sec=600)
        return renomear_e_converter(downloaded_file, "olist_relatorio_empresa2")

    def baixar_relatorio_2() -> Path:
        print("Iniciando download do Relatório 2...")
        fechar_modal_se_existir()
        abrir_menu_principal()
        clicar('//*[@id="main-menu"]/div[2]/div[1]/div[1]/nav/ul/li[4]/a/span')
        clicar('//*[@id="main-menu"]/div[2]/div[2]/nav[4]/ul/li[16]/a/span')
        clicar('//*[@id="root-relatorios-sistema"]/div/div[2]/div/div/div/div/a[2]/span[1]')
        clicar('//*[@id="root-relatorios-personalizados"]/div/div[2]/div[2]/div/div[2]/button')
        clicar('//*[@id="modal-filtros-relatorio-personalizado"]/div/div/div[2]/div/div[2]/div/div[4]/button')
        clicar('//*[@id="modal-filtros-relatorio-personalizado"]/div/div/div[3]/button[1]')

        before_files = snapshot_downloads()
        clicar('//*[@id="root-relatorios-personalizados"]/div/div[1]/div[1]/div/div/div[1]/button[1]')

        downloaded_file = esperar_download(before_files, timeout_sec=600)
        return renomear_e_converter(downloaded_file, "olist_relatorio2")

    try:
        fazer_login(username_1, password_1, "empresa 1")
        arquivo_1 = baixar_relatorio_1()
        arquivo_2 = baixar_relatorio_2()

        fazer_logout()

        fazer_login(username_2, password_2, "empresa 2")
        arquivo_3 = baixar_relatorio_1_empresa2()

        print("Os relatórios do Olist foram extraídos com sucesso.")
        return arquivo_1, arquivo_2, arquivo_3

    finally:
        driver.quit()


if __name__ == "__main__":
    run(headless=True)
