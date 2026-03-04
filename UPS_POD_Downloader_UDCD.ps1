# UPS_POD_Downloader.ps1
# UPS Proof of Delivery automatizált letöltő
# Futtatás: Jobb klikk -> Run with PowerShell

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "UPS POD Letöltő"
$form.Size = New-Object System.Drawing.Size(650, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(10, 10)
$headerLabel.Size = New-Object System.Drawing.Size(600, 30)
$headerLabel.Text = "UPS Proof of Delivery automatizált letöltő"
$headerLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = "DarkBlue"
$form.Controls.Add($headerLabel)

$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(10, 50)
$infoPanel.Size = New-Object System.Drawing.Size(600, 80)
$infoPanel.BorderStyle = "FixedSingle"
$infoPanel.BackColor = "LightYellow"

$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(10, 5)
$infoLabel.Size = New-Object System.Drawing.Size(580, 70)
$infoLabel.Text = "Használat:`n1. Add meg az UPS URL-t`n2. Válaszd ki az Excel fájlt és a letöltési mappát`n3. Add meg az UPS felhasználóneved és jelszavad`n4. Kattints a Letöltés indítása gombra"
$infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$infoPanel.Controls.Add($infoLabel)
$form.Controls.Add($infoPanel)

$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Location = New-Object System.Drawing.Point(10, 140)
$urlLabel.Size = New-Object System.Drawing.Size(120, 25)
$urlLabel.Text = "UPS URL:"
$urlLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($urlLabel)

$urlBox = New-Object System.Windows.Forms.TextBox
$urlBox.Location = New-Object System.Drawing.Point(140, 140)
$urlBox.Size = New-Object System.Drawing.Size(470, 25)
$urlBox.Text = "https://www.ups.com/track?loc=en_US&requester=ST/"
$urlBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($urlBox)

$excelLabel = New-Object System.Windows.Forms.Label
$excelLabel.Location = New-Object System.Drawing.Point(10, 180)
$excelLabel.Size = New-Object System.Drawing.Size(120, 25)
$excelLabel.Text = "Excel fájl:"
$excelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($excelLabel)

$excelBox = New-Object System.Windows.Forms.TextBox
$excelBox.Location = New-Object System.Drawing.Point(140, 180)
$excelBox.Size = New-Object System.Drawing.Size(370, 25)
$excelBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($excelBox)

$excelButton = New-Object System.Windows.Forms.Button
$excelButton.Location = New-Object System.Drawing.Point(520, 180)
$excelButton.Size = New-Object System.Drawing.Size(90, 25)
$excelButton.Text = "Tallózás"
$excelButton.Font = New-Object System.Drawing.Font("Arial", 9)
$excelButton.BackColor = "LightGray"
$excelButton.Add_Click({
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $fileBrowser.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
    $fileBrowser.Title = "Válaszd ki az Excel fájlt"
    if ($fileBrowser.ShowDialog() -eq "OK") { $excelBox.Text = $fileBrowser.FileName }
})
$form.Controls.Add($excelButton)

$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Location = New-Object System.Drawing.Point(10, 220)
$folderLabel.Size = New-Object System.Drawing.Size(120, 25)
$folderLabel.Text = "Letöltési mappa:"
$folderLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($folderLabel)

$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Location = New-Object System.Drawing.Point(140, 220)
$folderBox.Size = New-Object System.Drawing.Size(370, 25)
$folderBox.Font = New-Object System.Drawing.Font("Arial", 10)
$folderBox.Text = [Environment]::GetFolderPath("Desktop")
$form.Controls.Add($folderBox)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Location = New-Object System.Drawing.Point(520, 220)
$folderButton.Size = New-Object System.Drawing.Size(90, 25)
$folderButton.Text = "Tallózás"
$folderButton.Font = New-Object System.Drawing.Font("Arial", 9)
$folderButton.BackColor = "LightGray"
$folderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Válaszd ki a letöltési mappát"
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq "OK") { $folderBox.Text = $folderBrowser.SelectedPath }
})
$form.Controls.Add($folderButton)

$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Location = New-Object System.Drawing.Point(10, 260)
$userLabel.Size = New-Object System.Drawing.Size(120, 25)
$userLabel.Text = "UPS Username:"
$userLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($userLabel)

$userBox = New-Object System.Windows.Forms.TextBox
$userBox.Location = New-Object System.Drawing.Point(140, 260)
$userBox.Size = New-Object System.Drawing.Size(470, 25)
$userBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($userBox)

$passLabel = New-Object System.Windows.Forms.Label
$passLabel.Location = New-Object System.Drawing.Point(10, 300)
$passLabel.Size = New-Object System.Drawing.Size(120, 25)
$passLabel.Text = "UPS Password:"
$passLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($passLabel)

$passBox = New-Object System.Windows.Forms.TextBox
$passBox.Location = New-Object System.Drawing.Point(140, 300)
$passBox.Size = New-Object System.Drawing.Size(470, 25)
$passBox.PasswordChar = '*'
$passBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($passBox)

$checkLabel = New-Object System.Windows.Forms.Label
$checkLabel.Location = New-Object System.Drawing.Point(10, 340)
$checkLabel.Size = New-Object System.Drawing.Size(600, 20)
$checkLabel.Text = "Az Excel-ben szükséges oszlopok:"
$checkLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($checkLabel)

$checkList = New-Object System.Windows.Forms.ListBox
$checkList.Location = New-Object System.Drawing.Point(10, 360)
$checkList.Size = New-Object System.Drawing.Size(600, 60)
$checkList.Font = New-Object System.Drawing.Font("Arial", 9)
$checkList.Items.AddRange(@(
    "✓ 'Tracking Number' - a nyomkövetési szám",
    "✓ 'összefűz' - a letöltött fájl végső neve (ű-vel!)",
    "✓ 'Date', 'Carton No', 'MO' - csak szín ellenőrzés"
))
$checkList.Enabled = $false
$checkList.BackColor = "White"
$form.Controls.Add($checkList)

$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(10, 430)
$logLabel.Size = New-Object System.Drawing.Size(600, 20)
$logLabel.Text = "Folyamat napló:"
$logLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10, 450)
$logBox.Size = New-Object System.Drawing.Size(600, 100)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = "Black"
$logBox.ForeColor = "Lime"
$form.Controls.Add($logBox)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 560)
$progressBar.Size = New-Object System.Drawing.Size(280, 25)
$form.Controls.Add($progressBar)

$script:stopRequested = $false
$script:pythonProcess = $null

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(300, 560)
$stopButton.Size = New-Object System.Drawing.Size(90, 25)
$stopButton.Text = "STOP Megállítás"
$stopButton.BackColor = "Orange"
$stopButton.ForeColor = "White"
$stopButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$stopButton.Enabled = $false
$stopButton.Add_Click({
    $script:stopRequested = $true
    Write-Log "LEALLAS: Leállítás kérve..."
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
        Set-Content -Path $stopFilePath -Value "stop" -Force
        Write-Log "   Jelzőfájl létrehozva"
        Start-Sleep -Seconds 3
        if (!$script:pythonProcess.HasExited) {
            $script:pythonProcess.Kill()
            Write-Log "   Python folyamat leallitva (KILL)"
        } else {
            Write-Log "   Python folyamat rendben leallt"
        }
    } else {
        Write-Log "   Nincs futó Python folyamat"
    }
})
$form.Controls.Add($stopButton)

$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(400, 560)
$startButton.Size = New-Object System.Drawing.Size(110, 25)
$startButton.Text = "Letöltés indítása"
$startButton.BackColor = "ForestGreen"
$startButton.ForeColor = "White"
$startButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($startButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(520, 600)
$exitButton.Size = New-Object System.Drawing.Size(90, 25)
$exitButton.Text = "Kilépés"
$exitButton.BackColor = "DarkRed"
$exitButton.ForeColor = "White"
$exitButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$exitButton.Add_Click({
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
        Set-Content -Path $stopFilePath -Value "stop" -Force
        Start-Sleep -Seconds 2
        if (!$script:pythonProcess.HasExited) { $script:pythonProcess.Kill() }
    }
    $form.Close()
})
$form.Controls.Add($exitButton)

function Write-Log {
    param($Message)
    $logBox.AppendText($Message + "`r`n")
    $logBox.Refresh()
    Start-Sleep -Milliseconds 10
}

$startButton.Add_Click({
    $startButton.Enabled = $false
    $stopButton.Enabled = $true
    $script:stopRequested = $false

    $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
    if (Test-Path $stopFilePath) { Remove-Item $stopFilePath -Force }

    $url          = $urlBox.Text.Trim()
    $excelPath    = $excelBox.Text.Trim()
    $downloadFolder = $folderBox.Text.Trim()
    $username     = $userBox.Text.Trim()
    $password     = $passBox.Text.Trim()

    if (-not $url)                                        { [System.Windows.Forms.MessageBox]::Show("Add meg az UPS URL-t!","Hiba","OK","Error"); $startButton.Enabled=$true;$stopButton.Enabled=$false;return }
    if (-not $excelPath -or -not (Test-Path $excelPath)) { [System.Windows.Forms.MessageBox]::Show("Érvényes Excel fájlt kell kiválasztani!","Hiba","OK","Error"); $startButton.Enabled=$true;$stopButton.Enabled=$false;return }
    if (-not $downloadFolder -or -not (Test-Path $downloadFolder)) { [System.Windows.Forms.MessageBox]::Show("Érvényes letöltési mappát kell kiválasztani!","Hiba","OK","Error"); $startButton.Enabled=$true;$stopButton.Enabled=$false;return }
    if (-not $username) { [System.Windows.Forms.MessageBox]::Show("Add meg az UPS felhasználóneved!","Hiba","OK","Error"); $startButton.Enabled=$true;$stopButton.Enabled=$false;return }
    if (-not $password) { [System.Windows.Forms.MessageBox]::Show("Add meg az UPS jelszavad!","Hiba","OK","Error"); $startButton.Enabled=$true;$stopButton.Enabled=$false;return }

    Write-Log "==========================================="
    Write-Log "UPS POD Letöltő indítása"
    Write-Log "==========================================="
    Write-Log "Dátum: $(Get-Date)"
    Write-Log "Excel: $excelPath"
    Write-Log "Letöltési mappa: $downloadFolder"
    Write-Log "URL: $url"
    Write-Log "Felhasznalo: $username"
    Write-Log ""

    $pythonScript = @'
import sys
import pandas as pd
import time
import os
import random
import base64

# undetected-chromedriver: átmegy az Akamai bot detektáláson
import undetected_chromedriver as uc

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

STOP_FILE = os.path.join(os.environ['TEMP'], 'ups_pod_stop.txt')
GREEN_COLOR = '92D050'

def should_stop():
    return os.path.exists(STOP_FILE)

def log_message(msg):
    print(f"LOG: {msg}"); sys.stdout.flush()
def log_error(msg, details=""):
    print(f"LOG: [HIBA] {msg}")
    if details: print(f"LOG:   {details}")
    sys.stdout.flush()
def log_success(msg):
    print(f"LOG: [OK] {msg}"); sys.stdout.flush()
def log_step(step, msg):
    print(f"LOG:   [{step}] {msg}"); sys.stdout.flush()
def update_progress(current, total):
    print(f"PROGRESS: {current},{total}"); sys.stdout.flush()

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.2))

def human_click(driver, element):
    actions = ActionChains(driver)
    actions.move_to_element(element)
    time.sleep(random.uniform(0.3, 0.8))
    actions.click()
    actions.perform()

def handle_mfa_popup(driver):
    try:
        try:
            skip_btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.af-nextButton"))
            )
            human_click(driver, skip_btn)
            log_success("MFA popup kihagyva")
            time.sleep(2)
            return True
        except:
            pass
        try:
            skip_btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Skip for now')]"))
            )
            human_click(driver, skip_btn)
            log_success("MFA popup kihagyva (szoveg alapjan)")
            time.sleep(2)
            return True
        except:
            pass
        log_step("MFA", "Nincs MFA popup")
        return False
    except Exception as e:
        log_step("MFA", f"Hiba: {str(e)}")
        return False

def close_policy_popup(driver):
    try:
        popup = driver.find_elements(By.CSS_SELECTOR, "#ups-updateProfile-popup-container")
        if not popup:
            return
        not_now_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".ups-notNowButton"))
        )
        human_click(driver, not_now_btn)
        log_success("Policy popup bezarva")
        time.sleep(1)
    except Exception as e:
        log_step("Policy", f"Hiba: {str(e)}")

def close_chat_if_present(driver):
    try:
        chat = driver.find_elements(By.CSS_SELECTOR, "div.WACBotContainer")
        if not chat:
            return
        close_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACHeader__CloseAndRestartButton"))
        )
        human_click(driver, close_btn)
        time.sleep(1)
        try:
            yes_btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACConfirmModal__YesButton"))
            )
            human_click(driver, yes_btn)
        except:
            pass
        log_success("Chat bezarva")
        time.sleep(1)
    except Exception as e:
        log_step("Chat", f"Hiba: {str(e)}")

def accept_cookies(driver):
    try:
        for by, selector, description in [
            (By.ID, "onetrust-accept-btn-handler", "Allow All (banner)"),
            (By.ID, "onetrust-reject-all-handler", "Essential Only (banner)"),
            (By.ID, "onetrust-pc-btn-handler", "Cookie Settings (banner)"),
            (By.ID, "accept-recommended-btn-handler", "Allow All (big)"),
            (By.CSS_SELECTOR, ".save-preference-btn-handler", "Confirm Choices (big)"),
            (By.ID, "close-pc-btn-handler", "Close X (big)")
        ]:
            try:
                btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((by, selector)))
                human_click(driver, btn)
                log_success(f"Cookie kezelve: {description}")
                time.sleep(1)
                return True
            except:
                continue
        log_step("Cookie", "Nincs cookie ablak")
        return False
    except Exception as e:
        log_step("Cookie", f"Hiba: {str(e)}")
        return False

def login_if_needed(driver, username, password):
    try:
        sign_in_btn = None
        for selector in [
            "//a[contains(text(),'Sign in')]",
            "//a[contains(text(),'Log in')]",
            "//a[contains(@href,'/account/login')]",
            "//button[contains(text(),'Sign in')]",
        ]:
            try:
                sign_in_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                break
            except:
                continue

        if not sign_in_btn:
            log_step("Login", "Mar be van jelentkezve")
            return False

        log_step("Login", "Bejelentkezes szukseges...")
        human_click(driver, sign_in_btn)
        time.sleep(random.uniform(2, 3.5))

        username_field = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        human_click(driver, username_field)
        time.sleep(random.uniform(0.5, 1.0))
        username_field.clear()
        human_type(username_field, username)
        time.sleep(random.uniform(1.0, 2.0))

        continue_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "._button-login-id"))
        )
        human_click(driver, continue_btn)
        log_step("Login", "Continue gomb megnyomva")
        time.sleep(random.uniform(2.0, 3.5))

        password_field = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        human_click(driver, password_field)
        time.sleep(random.uniform(0.5, 1.0))
        password_field.clear()
        human_type(password_field, password)
        time.sleep(random.uniform(1.0, 2.0))

        login_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "._button-login-password"))
        )
        human_click(driver, login_btn)
        log_success("Bejelentkezes sikeres")
        time.sleep(random.uniform(3, 5))
        handle_mfa_popup(driver)
        return True

    except Exception as e:
        log_error("Bejelentkezesi hiba", str(e))
        return False

def is_row_processed(ws, row_idx):
    for col in range(1, 6):
        cell = ws.cell(row=row_idx, column=col)
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
            if cell.fill.fgColor.rgb[-6:] == GREEN_COLOR:
                return True
    return False

def save_pod_pdf(driver, download_folder, new_name, tracking_window):
    """
    PDF mentes a POD ablakbol.
    driver -> jelenleg a POD ablakra mutat
    tracking_window -> az eredeti tracking oldal ablak handle
    finally -> minden extra ablak bezarva, visszavaltas tracking_window-ra
    """
    try:
        windows_before = set(driver.window_handles)

        log_step("PDF", "Print gomb keresese a POD ablakban...")
        print_btn = None
        for by, sel, desc in [
            (By.ID, "stApp_POD_btnPrint", "ID: stApp_POD_btnPrint"),
            (By.LINK_TEXT, "Print this page", "Link szoveg"),
            (By.PARTIAL_LINK_TEXT, "Print", "Reszleges")
        ]:
            try:
                print_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((by, sel))
                )
                log_step("PDF", f"Print gomb talalva: {desc}")
                break
            except:
                continue

        if not print_btn:
            log_error("Print gomb nem talalhato a POD ablakban")
            return False

        human_click(driver, print_btn)
        log_success("Print gomb megnyomva")
        time.sleep(2)

        # Varj a nyomtatasi ablakra es valts at
        try:
            WebDriverWait(driver, 8).until(
                lambda d: len(d.window_handles) > len(windows_before)
            )
            windows_after = set(driver.window_handles)
            new_windows = windows_after - windows_before
            if new_windows:
                print_window = new_windows.pop()
                driver.switch_to.window(print_window)
                log_success("Nyomtatasi ablakra valtva")
                time.sleep(2)
            else:
                log_step("PDF", "Nincs uj nyomtatasi ablak, maradunk")
        except TimeoutException:
            log_step("PDF", "Nyomtatasi ablak nem nyilt, maradunk")

        # CDP PDF mentes
        log_step("PDF", "CDP PDF mentes...")
        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "paperWidth": 8.27,
            "paperHeight": 11.69,
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
        })

        pdf_bytes = base64.b64decode(pdf_data['data'])
        output_path = os.path.join(download_folder, f"{new_name}.pdf")
        if os.path.exists(output_path):
            os.remove(output_path)
        with open(output_path, 'wb') as f:
            f.write(pdf_bytes)
        log_success(f"PDF mentve: {new_name}.pdf ({len(pdf_bytes)} bytes)")
        return True

    except Exception as e:
        log_error("PDF mentes hiba", str(e))
        return False

    finally:
        # Bezarunk MINDEN ablakot kiveve az eredeti tracking ablakot
        try:
            for handle in list(driver.window_handles):
                if handle != tracking_window:
                    driver.switch_to.window(handle)
                    driver.close()
                    log_step("Ablak", "Extra ablak bezarva")
        except Exception as e:
            log_step("Ablak", f"Bezarasi hiba: {str(e)}")
        try:
            driver.switch_to.window(tracking_window)
            log_step("Ablak", "Visszavaltas tracking ablakra")
        except:
            if driver.window_handles:
                driver.switch_to.window(driver.window_handles[0])
                log_step("Ablak", "Visszavaltas elso ablakra")

def main():
    if len(sys.argv) < 6:
        log_error("Hianyzo argumentumok"); return 1
    ups_url         = sys.argv[1]
    excel_path      = sys.argv[2]
    download_folder = sys.argv[3]
    UPS_USERNAME    = sys.argv[4]
    UPS_PASSWORD    = sys.argv[5]

    log_message("="*60)
    log_message("PYTHON SCRIPT FUT - undetected-chromedriver mod")
    log_message("="*60)
    log_message(f"Excel: {excel_path}")
    log_message(f"Mappa: {download_folder}")
    log_message(f"URL: {ups_url}")
    log_message(f"Felhasznalo: {UPS_USERNAME}")
    log_message("")

    # --- Excel beolvasas ---
    log_message("[1/5] Excel fajl beolvasasa...")
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        log_success(f"Excel beolvasva - {len(df)} sor, oszlopok: {list(df.columns)}")
    except Exception as e:
        log_error("Excel olvasasi hiba", str(e)); return 1

    required = ['Tracking Number', 'összefűz']
    missing = [c for c in required if c not in df.columns]
    if missing:
        log_error("Hianyzó oszlopok", f"Kell: {required}, Hianyzik: {missing}"); return 1

    try:
        wb = load_workbook(excel_path)
        ws = wb.active
    except Exception as e:
        log_error("Excel megnyitasi hiba", str(e)); return 1

    to_process_indices = []
    for idx, row in df.iterrows():
        excel_row = idx + 2
        if is_row_processed(ws, excel_row):
            log_step("Szures", f"Sor {excel_row} mar feldolgozva (zold), kihagyva")
            continue
        tracking = str(row['Tracking Number']).strip() if pd.notna(row['Tracking Number']) else ''
        new_name = str(row['összefűz']).strip() if pd.notna(row['összefűz']) else ''
        if not tracking or not new_name:
            log_step("Szures", f"Sor {excel_row} hianyos, kihagyva")
            continue
        to_process_indices.append((idx, excel_row, tracking, new_name))

    total = len(to_process_indices)
    if total == 0:
        log_message("Nincs feldolgozando sor."); return 0
    log_success(f"Feldolgozando sorok: {total}")
    update_progress(0, total)
    log_message("")

    # --- Bongeszo inditas undetected-chromedriver-rel ---
    log_message("[2/5] Bongeszo inditasa (undetected-chromedriver)...")
    try:
        options = uc.ChromeOptions()
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--start-maximized")
        options.add_argument("--lang=hu-HU")

        # Letoltesi beallitasok
        prefs = {
            "download.default_directory": download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "credentials_enable_service": False
        }
        options.add_experimental_option("prefs", prefs)

        driver = uc.Chrome(options=options)
        log_success("Bongeszo sikeresen elindult (undetected mod)")

    except Exception as e:
        log_error("Bongeszo inditasi hiba", str(e))
        log_error("Telepitsd: pip install undetected-chromedriver")
        return 1

    try:
        driver.get(ups_url)
        time.sleep(4)
        log_success("Oldal betoltve")

        handle_mfa_popup(driver)
        accept_cookies(driver)
        login_if_needed(driver, UPS_USERNAME, UPS_PASSWORD)
        accept_cookies(driver)
        log_message("")

        processed    = 0
        success_count = 0
        zold_fill = PatternFill(start_color=GREEN_COLOR, end_color=GREEN_COLOR, fill_type='solid')

        for idx, excel_row, tracking, new_name in to_process_indices:
            if should_stop():
                log_message("Leallitasi keres eszlelve..."); break

            log_message("")
            log_message("-"*50)
            log_message(f"Feldolgozas: {tracking} -> {new_name} (Excel sor: {excel_row})")
            log_message("-"*50)

            # --- TRACKING SZAM BEIRASA ---
            log_step("3a", "Tracking szam mezo keresese...")
            track_input = None
            for by, sel, desc in [
                (By.ID, "stApp_trackingNumber", "ID: stApp_trackingNumber"),
                (By.CSS_SELECTOR, "textarea[formcontrolname='trackingNumber']", "Angular form control"),
                (By.CSS_SELECTOR, "textarea.ups-textbox_textarea", "Class"),
                (By.NAME, "trackingnumber", "NAME")
            ]:
                try:
                    el = WebDriverWait(driver, 3).until(EC.presence_of_element_located((by, sel)))
                    track_input = el
                    log_step("3a", f"Megtalalva: {desc}")
                    break
                except:
                    continue

            if not track_input:
                log_error("Tracking mezo nem talalhato"); continue

            # Kattintás a mezőre
            human_click(driver, track_input)
            time.sleep(random.uniform(0.5, 1.0))

            # Mező törlése
            track_input.clear()
            time.sleep(0.2)
            track_input.send_keys(Keys.CONTROL + "a")
            track_input.send_keys(Keys.DELETE)
            time.sleep(0.3)

            # Beírás JavaScript-tel (Angular input+change event triggerelése)
            driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                track_input, tracking
            )
            log_step("3a", f"Tracking szam beillesztve: '{tracking}'")
            time.sleep(random.uniform(0.5, 1.0))

            # Angular blur event - mező elhagyása
            track_input.send_keys(Keys.TAB)
            time.sleep(random.uniform(1.0, 1.5))

            # Ellenorzés
            try:
                actual_value = track_input.get_attribute('value')
                log_step("3a", f"Mezo tartalma: '{actual_value}'")
                if actual_value.strip() != tracking.strip():
                    log_step("3a", "Ertek nem egyezik, ujra probaljuk...")
                    human_click(driver, track_input)
                    track_input.clear()
                    time.sleep(0.5)
                    driver.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                        track_input, tracking
                    )
                    time.sleep(0.8)
                    track_input.send_keys(Keys.TAB)
                    time.sleep(0.8)
            except:
                pass

            # --- TRACK GOMB ---
            log_step("3b", "Track gomb keresese...")
            try:
                track_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "stApp_btnTrack"))
                )
                log_success("Track gomb megtalálva")
                human_click(driver, track_btn)
                log_success("Track gomb megnyomva")
                handle_mfa_popup(driver)

                # POD gombra varunk (5 masodpercenkent ellenorizzuk, max 30 mp)
                log_step("Varas", "POD gombra varunk (max 30mp)...")
                pod_found = False
                for attempt in range(6):  # 6 x 5mp = 30mp
                    time.sleep(5)
                    els = driver.find_elements(By.ID, "stApp_btnProofOfDeliveryonDetails")
                    if els:
                        log_success(f"POD gomb megjelent ({(attempt+1)*5}mp utan)")
                        pod_found = True
                        break
                    log_step("Varas", f"{(attempt+1)*5}mp... meg varakozunk")

                if not pod_found:
                    log_error(f"POD gomb 30mp utan sem jelent meg, URL: {driver.current_url}")
                    log_step("Retry", "Oldal frissitese, ujra probalkozas (5mp)...")
                    driver.refresh()
                    time.sleep(5)
                    els = driver.find_elements(By.ID, "stApp_btnProofOfDeliveryonDetails")
                    if els:
                        log_success("POD gomb megjelent frissites utan")
                        pod_found = True
                    else:
                        log_error("POD gomb frissites utan sem jelent meg, sor kihagyva")
                        continue

            except Exception as e:
                log_error("Track gomb hiba", str(e))
                continue

            close_policy_popup(driver)
            close_chat_if_present(driver)

            # --- POD LINK ---
            log_step("3c", "POD link keresese...")
            pod_link = None
            used = ""
            for by, sel, desc in [
                (By.ID, "stApp_btnProofOfDeliveryonDetails", "ID"),
                (By.LINK_TEXT, "Proof of Delivery", "Link szoveg"),
                (By.PARTIAL_LINK_TEXT, "Proof", "Reszleges")
            ]:
                try:
                    el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, sel)))
                    pod_link = el
                    used = desc
                    log_step("3c", f"POD link talalva: {desc}")
                    break
                except:
                    continue

            if not pod_link:
                log_error("POD link nem talalhato"); continue

            # Az eredeti tracking ablak handle
            tracking_window = driver.current_window_handle
            human_click(driver, pod_link)
            log_success(f"POD link megnyitva ({used})")

            # Varj a POD ablakra es valts at ra
            try:
                WebDriverWait(driver, 8).until(
                    lambda d: len(d.window_handles) > 1
                )
                for w in driver.window_handles:
                    if w != tracking_window:
                        driver.switch_to.window(w)
                        break
                log_success("POD ablakra valtva")
                time.sleep(3)
            except Exception as e:
                log_step("Ablak", f"POD ablak nem nyilt: {str(e)}")

            # PDF mentes a POD ablakbol
            pdf_saved = save_pod_pdf(driver, download_folder, new_name, tracking_window)

            if pdf_saved:
                for col in range(1, 6):
                    ws.cell(row=excel_row, column=col).fill = zold_fill
                log_success(f"Sor {excel_row} zoldre szinezve")
                success_count += 1
            else:
                log_error("PDF mentes sikertelen")

            # --- VISSZANAVIGALAS ---
            log_step("Nav", "Visszanavigalas a tracking foroldalra...")
            driver.get(ups_url)
            time.sleep(random.uniform(3, 5))

            handle_mfa_popup(driver)
            login_if_needed(driver, UPS_USERNAME, UPS_PASSWORD)

            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, "stApp_trackingNumber"))
                )
                log_success("Tracking oldal keszen all")
                time.sleep(random.uniform(1.5, 2.5))
            except TimeoutException:
                log_error("Tracking mezo nem jelent meg, folytatjuk...")

            processed += 1
            update_progress(processed, total)
            log_success(f"Feldolgozva: {processed}/{total}")

        # --- Excel mentes ---
        log_message("\n[4/5] Excel fajl mentese...")
        output_path = excel_path.replace('.xlsx', '_FELDOLGOZOTT.xlsx')
        if output_path == excel_path:
            output_path = excel_path + '_FELDOLGOZOTT.xlsx'
        try:
            wb.save(output_path)
            log_success(f"Excel mentve: {output_path}")
            log_message(f"Sikeres: {success_count}/{total}\n")
        except Exception as e:
            log_error("Excel mentesi hiba", str(e)); return 1

        log_message("[5/5] Folyamat befejezve")
        return 0

    except Exception as e:
        log_error("Varatlan hiba", str(e)); return 1
    finally:
        try:
            driver.quit()
            log_message("Bongeszo bezarva")
        except:
            pass
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)

if __name__ == "__main__":
    sys.exit(main())
'@

    $tempPython = [System.IO.Path]::GetTempFileName() + ".py"
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tempPython, $pythonScript, $utf8WithBom)

    Write-Log "Python script futtatasa..."
    Write-Log "(Elso futtatasnal az undetected-chromedriver letoltheti a Chrome drivert - ez normal)"
    Write-Log ""

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "python"
    $psi.Arguments = "`"$tempPython`" `"$url`" `"$excelPath`" `"$downloadFolder`" `"$username`" `"$password`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $script:pythonProcess = $process

    $outputEvent = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            if ($data.StartsWith("LOG: ")) {
                $message = $data.Substring(5)
                $form.BeginInvoke([Action]{ Write-Log $message })
            } elseif ($data.StartsWith("PROGRESS: ")) {
                $parts = $data.Substring(10).Split(',')
                if ($parts.Count -eq 2) {
                    $current = [int]$parts[0]; $total = [int]$parts[1]
                    $form.BeginInvoke([Action]{ $progressBar.Maximum = $total; $progressBar.Value = $current })
                }
            }
        }
    }

    $errorEvent = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            $form.BeginInvoke([Action]{ Write-Log "PYTHON HIBA: $data" })
            $hibaUzenet = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $data`r`n"
            Add-Content -Path "C:\temp\python_hibak.log" -Value $hibaUzenet
        }
    }

    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    $script:pythonProcess = $null

    Unregister-Event -SourceIdentifier $outputEvent.Name -Force -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $errorEvent.Name -Force -ErrorAction SilentlyContinue
    Remove-Item $tempPython -Force -ErrorAction SilentlyContinue

    Write-Log ""
    Write-Log "="*50
    if ($exitCode -eq 0) {
        Write-Log "FOLYAMAT SIKERESEN BEFEJEZODOTT"
        [System.Windows.Forms.MessageBox]::Show("A letöltés sikeresen befejeződött!", "Siker", "OK", "Information")
    } else {
        Write-Log "HIBA TORTENT (kód: $exitCode)"
        [System.Windows.Forms.MessageBox]::Show("Hiba történt! Ellenőrizd a naplót.", "Hiba", "OK", "Error")
    }
    Write-Log "="*50

    $progressBar.Value = 0
    $startButton.Enabled = $true
    $stopButton.Enabled = $false
})

$form.ShowDialog()
