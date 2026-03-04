# UPS_POD_Downloader_Selenium_v1.ps1
# UPS Proof of Delivery automatizált letöltő
# Futtatás: Jobb klikk -> Run with PowerShell

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "UPS POD Letöltő"
$form.Size = New-Object System.Drawing.Size(650, 820)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

# --- Fejléc ---
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(10, 10)
$headerLabel.Size = New-Object System.Drawing.Size(600, 30)
$headerLabel.Text = "UPS Proof of Delivery automatizált letöltő"
$headerLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = "DarkBlue"
$form.Controls.Add($headerLabel)

# --- Használati útmutató panel ---
$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(10, 50)
$infoPanel.Size = New-Object System.Drawing.Size(600, 100)
$infoPanel.BorderStyle = "FixedSingle"
$infoPanel.BackColor = "LightYellow"

$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(10, 5)
$infoLabel.Size = New-Object System.Drawing.Size(580, 90)
$infoLabel.Text = "Használat:`n1. Kattints a 'POD Chrome indítása' gombra - megnyílik egy Chrome ablak`n2. Jelentkezz be az UPS-be ebben a Chrome-ban (csak egyszer kell!)`n3. Válaszd ki az Excel fájlt és a letöltési mappát`n4. Kattints a 'Letöltés indítása' gombra"
$infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$infoPanel.Controls.Add($infoLabel)
$form.Controls.Add($infoPanel)

# --- POD Chrome indítása gomb ---
$chromePanelLabel = New-Object System.Windows.Forms.Label
$chromePanelLabel.Location = New-Object System.Drawing.Point(10, 162)
$chromePanelLabel.Size = New-Object System.Drawing.Size(600, 20)
$chromePanelLabel.Text = "1. lépés: POD Chrome indítása"
$chromePanelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$chromePanelLabel.ForeColor = "DarkBlue"
$form.Controls.Add($chromePanelLabel)

$launchChromeButton = New-Object System.Windows.Forms.Button
$launchChromeButton.Location = New-Object System.Drawing.Point(10, 185)
$launchChromeButton.Size = New-Object System.Drawing.Size(220, 35)
$launchChromeButton.Text = "POD Chrome indítása"
$launchChromeButton.BackColor = "SteelBlue"
$launchChromeButton.ForeColor = "White"
$launchChromeButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$launchChromeButton.Add_Click({
    # Leállítjuk ha már fut egy ilyen Chrome
    $existing = Get-Process chrome -ErrorAction SilentlyContinue | Where-Object {
        $_.CommandLine -like "*SeleniumProfile*" -or $_.CommandLine -like "*9222*"
    }
    if ($existing) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Már fut egy POD Chrome. Újraindítod?", "POD Chrome", "YesNo", "Question")
        if ($result -eq "Yes") {
            $existing | Stop-Process -Force
            Start-Sleep -Seconds 2
        } else {
            Write-Log "POD Chrome már fut, folytatjuk..."
            return
        }
    }

    # Chrome elérési út keresése
    $chromePaths = @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $chromePath = $null
    foreach ($p in $chromePaths) {
        if (Test-Path $p) { $chromePath = $p; break }
    }
    if (-not $chromePath) {
        [System.Windows.Forms.MessageBox]::Show(
            "A Chrome nem található! Keresd meg manuálisan.", "Hiba", "OK", "Error")
        return
    }

    $profileDir = "C:\SeleniumProfile"
    Start-Process $chromePath -ArgumentList "--remote-debugging-port=9222 --user-data-dir=`"$profileDir`""
    Write-Log "POD Chrome elindítva (debug port: 9222, profil: $profileDir)"
    Write-Log ">>> Jelentkezz be az UPS-be a megnyílt Chrome-ban, majd kattints a Letöltés indítása gombra!"
    $chromeStatusLabel.Text = "✓ POD Chrome fut - jelentkezz be az UPS-be!"
    $chromeStatusLabel.ForeColor = "DarkGreen"
})
$form.Controls.Add($launchChromeButton)

$chromeStatusLabel = New-Object System.Windows.Forms.Label
$chromeStatusLabel.Location = New-Object System.Drawing.Point(240, 193)
$chromeStatusLabel.Size = New-Object System.Drawing.Size(380, 20)
$chromeStatusLabel.Text = "Chrome még nem indult el"
$chromeStatusLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$chromeStatusLabel.ForeColor = "Gray"
$form.Controls.Add($chromeStatusLabel)

# --- Elválasztó ---
$sep1 = New-Object System.Windows.Forms.Label
$sep1.Location = New-Object System.Drawing.Point(10, 228)
$sep1.Size = New-Object System.Drawing.Size(600, 2)
$sep1.BorderStyle = "Fixed3D"
$form.Controls.Add($sep1)

# --- 2. lépés fejléc ---
$step2Label = New-Object System.Windows.Forms.Label
$step2Label.Location = New-Object System.Drawing.Point(10, 235)
$step2Label.Size = New-Object System.Drawing.Size(600, 20)
$step2Label.Text = "2. lépés: Fájlok és beállítások"
$step2Label.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$step2Label.ForeColor = "DarkBlue"
$form.Controls.Add($step2Label)

# --- Excel fájl ---
$excelLabel = New-Object System.Windows.Forms.Label
$excelLabel.Location = New-Object System.Drawing.Point(10, 262)
$excelLabel.Size = New-Object System.Drawing.Size(120, 25)
$excelLabel.Text = "Excel fájl:"
$excelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($excelLabel)

$excelBox = New-Object System.Windows.Forms.TextBox
$excelBox.Location = New-Object System.Drawing.Point(140, 262)
$excelBox.Size = New-Object System.Drawing.Size(370, 25)
$excelBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($excelBox)

$excelButton = New-Object System.Windows.Forms.Button
$excelButton.Location = New-Object System.Drawing.Point(520, 262)
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

# --- Letöltési mappa ---
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Location = New-Object System.Drawing.Point(10, 300)
$folderLabel.Size = New-Object System.Drawing.Size(120, 25)
$folderLabel.Text = "Letöltési mappa:"
$folderLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($folderLabel)

$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Location = New-Object System.Drawing.Point(140, 300)
$folderBox.Size = New-Object System.Drawing.Size(370, 25)
$folderBox.Font = New-Object System.Drawing.Font("Arial", 10)
$folderBox.Text = [Environment]::GetFolderPath("Desktop")
$form.Controls.Add($folderBox)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Location = New-Object System.Drawing.Point(520, 300)
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

# --- UPS URL ---
$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Location = New-Object System.Drawing.Point(10, 338)
$urlLabel.Size = New-Object System.Drawing.Size(120, 25)
$urlLabel.Text = "UPS URL:"
$urlLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($urlLabel)

$urlBox = New-Object System.Windows.Forms.TextBox
$urlBox.Location = New-Object System.Drawing.Point(140, 338)
$urlBox.Size = New-Object System.Drawing.Size(470, 25)
$urlBox.Text = "https://www.ups.com/track?loc=en_US&requester=ST/"
$urlBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($urlBox)

# --- Excel oszlopok info ---
$checkLabel = New-Object System.Windows.Forms.Label
$checkLabel.Location = New-Object System.Drawing.Point(10, 375)
$checkLabel.Size = New-Object System.Drawing.Size(600, 20)
$checkLabel.Text = "Az Excel-ben szükséges oszlopok:"
$checkLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($checkLabel)

$checkList = New-Object System.Windows.Forms.ListBox
$checkList.Location = New-Object System.Drawing.Point(10, 395)
$checkList.Size = New-Object System.Drawing.Size(600, 50)
$checkList.Font = New-Object System.Drawing.Font("Arial", 9)
$checkList.Items.AddRange(@(
    "✓ 'Tracking Number' - a nyomkövetési szám",
    "✓ 'összefűz' - a letöltött fájl végső neve (ű-vel!)"
))
$checkList.Enabled = $false
$checkList.BackColor = "White"
$form.Controls.Add($checkList)

# --- Napló ---
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(10, 455)
$logLabel.Size = New-Object System.Drawing.Size(600, 20)
$logLabel.Text = "Folyamat napló:"
$logLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10, 475)
$logBox.Size = New-Object System.Drawing.Size(600, 150)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = "Black"
$logBox.ForeColor = "Lime"
$form.Controls.Add($logBox)

# --- Progress bar ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 635)
$progressBar.Size = New-Object System.Drawing.Size(380, 25)
$form.Controls.Add($progressBar)

# --- STOP gomb ---
$script:stopRequested = $false
$script:pythonProcess = $null

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(400, 635)
$stopButton.Size = New-Object System.Drawing.Size(90, 25)
$stopButton.Text = "STOP"
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
        Start-Sleep -Seconds 3
        if (!$script:pythonProcess.HasExited) {
            $script:pythonProcess.Kill()
            Write-Log "   Python folyamat leallitva (KILL)"
        }
    }
})
$form.Controls.Add($stopButton)

# --- Letöltés indítása gomb ---
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(500, 635)
$startButton.Size = New-Object System.Drawing.Size(120, 25)
$startButton.Text = "Letöltés indítása"
$startButton.BackColor = "ForestGreen"
$startButton.ForeColor = "White"
$startButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($startButton)

# --- Kilépés gomb ---
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(520, 670)
$exitButton.Size = New-Object System.Drawing.Size(100, 25)
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
    $logBox.ScrollToCaret()
    $logBox.Refresh()
    Start-Sleep -Milliseconds 10
}

# =====================================================
# LETÖLTÉS INDÍTÁSA
# =====================================================
$startButton.Add_Click({
    $startButton.Enabled = $false
    $stopButton.Enabled = $true
    $script:stopRequested = $false

    $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
    if (Test-Path $stopFilePath) { Remove-Item $stopFilePath -Force }

    $url            = $urlBox.Text.Trim()
    $excelPath      = $excelBox.Text.Trim()
    $downloadFolder = $folderBox.Text.Trim()

    if (-not $url) {
        [System.Windows.Forms.MessageBox]::Show("Add meg az UPS URL-t!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }
    if (-not $excelPath -or -not (Test-Path $excelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes Excel fájlt kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }
    if (-not $downloadFolder -or -not (Test-Path $downloadFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes letöltési mappát kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }

    # Ellenőrzés: fut-e a POD Chrome debug porton
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:9222/json" -TimeoutSec 2 -ErrorAction Stop
        Write-Log "POD Chrome detektálva a 9222-es porton - OK"
    } catch {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "A POD Chrome nem fut vagy nem válaszol a 9222-es porton!`n`nElőször kattints a 'POD Chrome indítása' gombra és jelentkezz be az UPS-be.",
            "POD Chrome nem fut", "OK", "Warning")
        $startButton.Enabled = $true; $stopButton.Enabled = $false; return
    }

    Write-Log "==========================================="
    Write-Log "UPS POD Letöltő indítása"
    Write-Log "==========================================="
    Write-Log "Dátum: $(Get-Date)"
    Write-Log "Excel: $excelPath"
    Write-Log "Letöltési mappa: $downloadFolder"
    Write-Log "URL: $url"
    Write-Log ""

    $pythonScript = @'
import sys
import pandas as pd
import time
import os
import random
import base64
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
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
    if len(sys.argv) < 4:
        log_error("Hianyzo argumentumok"); return 1
    ups_url         = sys.argv[1]
    excel_path      = sys.argv[2]
    download_folder = sys.argv[3]

    log_message("="*60)
    log_message("PYTHON SCRIPT FUT - debuggerAddress mod")
    log_message("="*60)
    log_message(f"Excel: {excel_path}")
    log_message(f"Mappa: {download_folder}")
    log_message(f"URL: {ups_url}")
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

    # --- Csatlakozas a mar futo Chrome-hoz ---
    log_message("[2/5] Csatlakozas a POD Chrome-hoz (port 9222)...")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        driver = webdriver.Chrome(options=chrome_options)
        log_success("Sikeresen csatlakozva a POD Chrome-hoz!")
        log_success(f"Jelenlegi URL: {driver.current_url}")
    except Exception as e:
        log_error("Csatlakozasi hiba", str(e))
        log_error("Biztositsd hogy a POD Chrome fut es be vagy jelentkezve az UPS-be!")
        return 1

    try:
        # Navigalas az UPS tracking oldalra
        log_step("Nav", f"Navigalas: {ups_url}")
        driver.get(ups_url)
        time.sleep(3)
        log_success("UPS tracking oldal betoltve")
        log_message("")

        processed     = 0
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
                    el = WebDriverWait(driver, 5).until(EC.presence_of_element_located((by, sel)))
                    track_input = el
                    log_step("3a", f"Megtalalva: {desc}")
                    break
                except:
                    continue

            if not track_input:
                log_error("Tracking mezo nem talalhato"); continue

            # Kattintas es torles
            human_click(driver, track_input)
            time.sleep(random.uniform(0.5, 1.0))
            track_input.clear()
            time.sleep(0.2)
            track_input.send_keys(Keys.CONTROL + "a")
            track_input.send_keys(Keys.DELETE)
            time.sleep(0.3)

            # Beiratas JavaScript-tel (Angular input+change event)
            driver.execute_script(
                "arguments[0].value = arguments[1];"
                "arguments[0].dispatchEvent(new Event('input', {bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change', {bubbles:true}));",
                track_input, tracking
            )
            log_step("3a", f"Tracking szam beillesztve: '{tracking}'")
            time.sleep(random.uniform(0.5, 1.0))

            # Angular blur event
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

                # POD gombra varunk (5 masodpercenkent, max 30 mp)
                log_step("Varas", "POD gombra varunk (max 30mp)...")
                pod_found = False
                for attempt in range(6):
                    time.sleep(5)
                    els = driver.find_elements(By.ID, "stApp_btnProofOfDeliveryonDetails")
                    if els:
                        log_success(f"POD gomb megjelent ({(attempt+1)*5}mp utan)")
                        pod_found = True
                        break
                    log_step("Varas", f"{(attempt+1)*5}mp... meg varakozunk")

                if not pod_found:
                    log_error(f"POD gomb 30mp utan sem jelent meg, URL: {driver.current_url}")
                    log_step("Retry", "Oldal frissitese (5mp)...")
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

            tracking_window = driver.current_window_handle
            human_click(driver, pod_link)
            log_success(f"POD link megnyitva ({used})")

            # Varj a POD ablakra es valts at
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

            # PDF mentes
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
        # Ne zarjuk be a Chrome-ot - a felhasznalo manuálisan zarhatja
        log_message("Script befejezve. A POD Chrome nyitva maradt.")
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)

if __name__ == "__main__":
    sys.exit(main())
'@

    $tempPython = [System.IO.Path]::GetTempFileName() + ".py"
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tempPython, $pythonScript, $utf8WithBom)

    Write-Log "Python script futtatasa..."
    Write-Log ""

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "python"
    $psi.Arguments = "`"$tempPython`" `"$url`" `"$excelPath`" `"$downloadFolder`""
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
            Add-Content -Path "C:\temp\python_hibak.log" -Value $hibaUzenet -ErrorAction SilentlyContinue
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
