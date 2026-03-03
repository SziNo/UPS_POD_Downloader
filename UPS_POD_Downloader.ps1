# UPS_POD_Downloader.ps1
# UPS Proof of Delivery automatizált letöltő
# Futtatás: Jobb klikk -> Run with PowerShell

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUI létrehozása
$form = New-Object System.Windows.Forms.Form
$form.Text = "UPS POD Letöltő"
$form.Size = New-Object System.Drawing.Size(650, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = "White"

# Címsor
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.Location = New-Object System.Drawing.Point(10, 10)
$headerLabel.Size = New-Object System.Drawing.Size(600, 30)
$headerLabel.Text = "UPS Proof of Delivery automatizált letöltő"
$headerLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$headerLabel.ForeColor = "DarkBlue"
$form.Controls.Add($headerLabel)

# Információs panel
$infoPanel = New-Object System.Windows.Forms.Panel
$infoPanel.Location = New-Object System.Drawing.Point(10, 50)
$infoPanel.Size = New-Object System.Drawing.Size(600, 100)
$infoPanel.BorderStyle = "FixedSingle"
$infoPanel.BackColor = "LightYellow"

$infoLabel = New-Object System.Windows.Forms.Label
$infoLabel.Location = New-Object System.Drawing.Point(10, 5)
$infoLabel.Size = New-Object System.Drawing.Size(580, 90)
$infoLabel.Text = "Használat:`n" +
                  "1. Jelentkezz be az UPS fiókodba a böngészőben`n" +
                  "2. Másold ki azt az URL-t, ahol a tracking number mező van`n" +
                  "3. Tallózással válaszd ki az Excel fájlt`n" +
                  "4. Tallózással válaszd ki a letöltési mappát"
$infoLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$infoPanel.Controls.Add($infoLabel)
$form.Controls.Add($infoPanel)

# UPS URL
$urlLabel = New-Object System.Windows.Forms.Label
$urlLabel.Location = New-Object System.Drawing.Point(10, 160)
$urlLabel.Size = New-Object System.Drawing.Size(120, 25)
$urlLabel.Text = "UPS URL:"
$urlLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($urlLabel)

$urlBox = New-Object System.Windows.Forms.TextBox
$urlBox.Location = New-Object System.Drawing.Point(140, 160)
$urlBox.Size = New-Object System.Drawing.Size(470, 25)
$urlBox.Text = "https://www.ups.com/track?loc=en_US"
$urlBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($urlBox)

# Excel fájl
$excelLabel = New-Object System.Windows.Forms.Label
$excelLabel.Location = New-Object System.Drawing.Point(10, 200)
$excelLabel.Size = New-Object System.Drawing.Size(120, 25)
$excelLabel.Text = "Excel fájl:"
$excelLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($excelLabel)

$excelBox = New-Object System.Windows.Forms.TextBox
$excelBox.Location = New-Object System.Drawing.Point(140, 200)
$excelBox.Size = New-Object System.Drawing.Size(370, 25)
$excelBox.Font = New-Object System.Drawing.Font("Arial", 10)
$form.Controls.Add($excelBox)

$excelButton = New-Object System.Windows.Forms.Button
$excelButton.Location = New-Object System.Drawing.Point(520, 200)
$excelButton.Size = New-Object System.Drawing.Size(90, 25)
$excelButton.Text = "Tallózás"
$excelButton.Font = New-Object System.Drawing.Font("Arial", 9)
$excelButton.BackColor = "LightGray"
$excelButton.Add_Click({
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $fileBrowser.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
    $fileBrowser.Title = "Válaszd ki az Excel fájlt"
    if ($fileBrowser.ShowDialog() -eq "OK") {
        $excelBox.Text = $fileBrowser.FileName
    }
})
$form.Controls.Add($excelButton)

# Letöltési mappa
$folderLabel = New-Object System.Windows.Forms.Label
$folderLabel.Location = New-Object System.Drawing.Point(10, 240)
$folderLabel.Size = New-Object System.Drawing.Size(120, 25)
$folderLabel.Text = "Letöltési mappa:"
$folderLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($folderLabel)

$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Location = New-Object System.Drawing.Point(140, 240)
$folderBox.Size = New-Object System.Drawing.Size(370, 25)
$folderBox.Font = New-Object System.Drawing.Font("Arial", 10)
$folderBox.Text = [Environment]::GetFolderPath("Desktop")
$form.Controls.Add($folderBox)

$folderButton = New-Object System.Windows.Forms.Button
$folderButton.Location = New-Object System.Drawing.Point(520, 240)
$folderButton.Size = New-Object System.Drawing.Size(90, 25)
$folderButton.Text = "Tallózás"
$folderButton.Font = New-Object System.Drawing.Font("Arial", 9)
$folderButton.BackColor = "LightGray"
$folderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Válaszd ki a letöltési mappát"
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $folderBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($folderButton)

# Ellenőrző lista
$checkLabel = New-Object System.Windows.Forms.Label
$checkLabel.Location = New-Object System.Drawing.Point(10, 280)
$checkLabel.Size = New-Object System.Drawing.Size(600, 20)
$checkLabel.Text = "Az Excel-ben szükséges oszlopok:"
$checkLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($checkLabel)

$checkList = New-Object System.Windows.Forms.ListBox
$checkList.Location = New-Object System.Drawing.Point(10, 300)
$checkList.Size = New-Object System.Drawing.Size(600, 60)
$checkList.Font = New-Object System.Drawing.Font("Arial", 9)
$checkList.Items.AddRange(@(
    "✓ 'Tracking Number' - a nyomkövetési szám",
    "✓ 'összefűz' - a letöltött fájl végső neve (ű-vel!)",
    "✓ 'Date', 'Carton No', 'MO' - ezeket a program nem módosítja, csak ellenőrzi a színüket"
))
$checkList.Enabled = $false
$checkList.BackColor = "White"
$form.Controls.Add($checkList)

# Naplózó mező
$logLabel = New-Object System.Windows.Forms.Label
$logLabel.Location = New-Object System.Drawing.Point(10, 380)
$logLabel.Size = New-Object System.Drawing.Size(600, 20)
$logLabel.Text = "Folyamat napló:"
$logLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($logLabel)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10, 400)
$logBox.Size = New-Object System.Drawing.Size(600, 120)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$logBox.BackColor = "Black"
$logBox.ForeColor = "Lime"
$form.Controls.Add($logBox)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 540)
$progressBar.Size = New-Object System.Drawing.Size(280, 25)
$form.Controls.Add($progressBar)

# ============================================
# STOP GOMB
# ============================================
$script:stopRequested = $false
$script:pythonProcess = $null

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(300, 540)
$stopButton.Size = New-Object System.Drawing.Size(90, 25)
$stopButton.Text = "STOP Megállítás"
$stopButton.BackColor = "Orange"
$stopButton.ForeColor = "White"
$stopButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$stopButton.Enabled = $false
$stopButton.Add_Click({
    $script:stopRequested = $true
    Write-Log "LEALLAS: Leállítás kérve... (a következő tracking után leáll)"
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
        Set-Content -Path $stopFilePath -Value "stop" -Force
        Write-Log "   Jelzőfájl létrehozva: $stopFilePath"
    }
})
$form.Controls.Add($stopButton)

# Indítás gomb
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(400, 540)
$startButton.Size = New-Object System.Drawing.Size(110, 25)
$startButton.Text = "Letöltés indítása"
$startButton.BackColor = "ForestGreen"
$startButton.ForeColor = "White"
$startButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($startButton)

function Write-Log {
    param($Message)
    $logBox.AppendText($Message + "`r`n")
    $logBox.Refresh()
    Start-Sleep -Milliseconds 10
}

# Indítás gomb eseménykezelő
$startButton.Add_Click({
    $startButton.Enabled = $false
    $stopButton.Enabled = $true
    $script:stopRequested = $false
    
    $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
    if (Test-Path $stopFilePath) { Remove-Item $stopFilePath -Force }
    
    $url = $urlBox.Text.Trim()
    $excelPath = $excelBox.Text.Trim()
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
    
    Write-Log "==========================================="
    Write-Log "UPS POD Letöltő indítása"
    Write-Log "==========================================="
    Write-Log "Dátum: $(Get-Date)"
    Write-Log "Excel: $excelPath"
    Write-Log "Letöltési mappa: $downloadFolder"
    Write-Log "UPS URL: $url"
    Write-Log ""
    
    # Python script – webdriver-manager, cookie kezelés, bejelentkezés
    $pythonScript = @'
import sys
import pandas as pd
import time
import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
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

def check_element(driver, by, selector, timeout=5, description=""):
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        log_step("Kereses", f"Megtalalva: {description} ({selector})")
        return element
    except TimeoutException:
        log_error(f"Nem talalhato: {description}", f"Selector: {selector}, idotullepes: {timeout}s")
        return None
    except Exception as e:
        log_error(f"Hiba a kereseskor: {description}", str(e))
        return None

def close_chat_if_present(driver):
    try:
        chat = driver.find_elements(By.CSS_SELECTOR, "div.WACBotContainer")
        if not chat:
            return
        log_step("Chat", "UPS Assistant chat eszlelve, bezaras...")
        close_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACHeader__CloseAndRestartButton"))
        )
        close_btn.click()
        time.sleep(1)
        yes_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.WACConfirmModal__YesButton"))
        )
        yes_btn.click()
        log_success("Chat bezarva")
        time.sleep(1)
    except Exception as e:
        log_step("Chat", f"Nem sikerult bezarni a chatet: {str(e)}")

def handle_chrome_print(driver):
    try:
        main = driver.current_window_handle
        for handle in driver.window_handles:
            if handle != main:
                driver.switch_to.window(handle)
                break
        log_step("Print", "Print ablakba valtottunk")
        time.sleep(2)

        script = """
        const printBtn = document.querySelector('cr-button.action-button');
        if (printBtn) {
            printBtn.click();
            return true;
        }
        return false;
        """
        clicked = driver.execute_script(script)
        if clicked:
            log_success("Print gomb megnyomva (shadow DOM)")
        else:
            log_error("Print gomb nem talalhato shadow DOM-ban")
        time.sleep(2)
        driver.switch_to.window(main)
    except Exception as e:
        log_error("Hiba a print ablak kezelesekor", str(e))

def accept_cookies(driver):
    """Cookie-k automatikus elfogadása többféle nyelvi verzióban."""
    try:
        cookie_selectors = [
            "//button[contains(text(),'Cookie')]",
            "//button[contains(text(),'Elfogad')]",
            "//button[contains(text(),'Accept')]",
            "//button[contains(text(),'Süti')]",
            "//button[contains(@class,'cookie')]",
            "//button[contains(@id,'cookie')]",
            "[id*='cookie'] button",
            "[class*='cookie'] button"
        ]
        
        for selector in cookie_selectors:
            try:
                cookie_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                cookie_btn.click()
                log_success("Cookie-k elfogadva")
                time.sleep(1)
                return True
            except:
                continue
        
        log_step("Cookie", "Nincs cookie elfogado ablak vagy mar elfogadva")
        return False
    except Exception as e:
        log_step("Cookie", f"Cookie kezelesi hiba (nem kritikus): {str(e)}")
        return False

def login_if_needed(driver):
    """
    Bejelentkezés automatizálása, ha szükséges.
    IDE ÍRD BE A TESZTELÉSHEZ A FELHASZNÁLÓNEVED ÉS JELSZAVAD!
    """
    try:
        # Ellenőrizzük, hogy van-e "Sign in" vagy "Bejelentkezés" gomb
        sign_in_selectors = [
            "//a[contains(text(),'Sign in')]",
            "//a[contains(text(),'Bejelentkezés')]",
            "//a[contains(@href,'/account/login')]",
            "//button[contains(text(),'Sign in')]"
        ]
        
        sign_in_btn = None
        for selector in sign_in_selectors:
            try:
                sign_in_btn = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                break
            except:
                continue
        
        if sign_in_btn:
            log_step("Login", "Bejelentkezes szukseges...")
            sign_in_btn.click()
            time.sleep(2)
            
            # =====================================================================
            # !!! IDE ÍRD BE A SAJÁT UPS FELHASZNÁLÓNEVED ÉS JELSZAVAD !!!
            # =====================================================================
            UPS_USERNAME = "DHLSC2022"   # <-- Ezt cseréld ki
            UPS_PASSWORD = "APIconnect5483167"          # <-- Ezt cseréld ki
            # =====================================================================
            
            # Felhasználónév mező
            username_field = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "email"))
            )
            username_field.clear()
            username_field.send_keys(UPS_USERNAME)
            log_step("Login", "Felhasznalonev megadva")
            
            # Jelszó mező
            password_field = driver.find_element(By.ID, "pwd")
            password_field.clear()
            password_field.send_keys(UPS_PASSWORD)
            log_step("Login", "Jelszo megadva")
            
            # Bejelentkezés gomb
            login_btn = driver.find_element(By.ID, "submitBtn")
            login_btn.click()
            
            log_success("Bejelentkezes sikeres")
            time.sleep(3)
            return True
        else:
            log_step("Login", "Mar be van jelentkezve")
            return False
    except Exception as e:
        log_error("Bejelentkezesi hiba (nem kritikus, lehet mar be van jelentkezve)", str(e))
        return False

def is_row_processed(ws, row_idx):
    for col in range(1, 6):
        cell = ws.cell(row=row_idx, column=col)
        if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
            color = cell.fill.fgColor.rgb[-6:]
            if color == GREEN_COLOR:
                return True
    return False

def main():
    if len(sys.argv) < 4:
        log_error("Hianyzo argumentumok"); return 1
    ups_url = sys.argv[1]
    excel_path = sys.argv[2]
    download_folder = sys.argv[3]

    log_message("="*60)
    log_message("PYTHON SCRIPT FUT")
    log_message("="*60)
    log_message(f"Excel: {excel_path}")
    log_message(f"Mappa: {download_folder}")
    log_message(f"URL: {ups_url}\n")

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
        log_error("Excel megnyitasi hiba (openpyxl)", str(e)); return 1

    to_process_indices = []
    for idx, row in df.iterrows():
        excel_row = idx + 2
        
        if is_row_processed(ws, excel_row):
            log_step("Szures", f"Sor {excel_row} mar fel van dolgozva (zold), kihagyva")
            continue
        
        tracking = str(row['Tracking Number']).strip() if pd.notna(row['Tracking Number']) else ''
        new_name = str(row['összefűz']).strip() if pd.notna(row['összefűz']) else ''
        
        if not tracking or not new_name:
            log_step("Szures", f"Sor {excel_row} hianyos (nincs Tracking Number vagy összefűz), kihagyva")
            continue
        
        to_process_indices.append((idx, excel_row, tracking, new_name))
    
    total = len(to_process_indices)
    if total == 0:
        log_message("Nincs feldolgozando sor."); return 0
    log_success(f"Feldolgozando sorok: {total}")
    update_progress(0, total)
    log_message("")

    log_message("[2/5] Böngésző indítása...")
    chrome_options = Options()
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        log_success("Bongeszo sikeresen elindult")
    except Exception as e:
        log_error("Bongeszo inditasi hiba", str(e)); return 1

    try:
        driver.get(ups_url)
        time.sleep(3)
        log_success("Oldal betoltve")
        
        # COOKIE-K ELFOGADÁSA
        accept_cookies(driver)
        
        # BEJELENTKEZÉS HA KELL
        login_if_needed(driver)
        
        # Ha esetleg a bejelentkezés után újra kell fogadni a cookie-kat
        accept_cookies(driver)
        
        log_message("")

        processed = 0
        success_count = 0
        zold_fill = PatternFill(start_color=GREEN_COLOR, end_color=GREEN_COLOR, fill_type='solid')

        for idx, excel_row, tracking, new_name in to_process_indices:
            if should_stop():
                log_message("Leallitasi keres eszlelve..."); break

            log_message("")
            log_message("-"*50)
            log_message(f"Feldolgozas: {tracking} -> {new_name} (Excel sor: {excel_row})")
            log_message("-"*50)

            log_step("3a", "Tracking szám mező keresése...")
            track_selectors = [
                (By.ID, "stApp_trackingNumber", "ID: stApp_trackingNumber"),
                (By.CSS_SELECTOR, "textarea[formcontrolname='trackingNumber']", "Angular form control"),
                (By.CSS_SELECTOR, "textarea.ups-textbox_textarea", "Class"),
                (By.NAME, "trackingnumber", "NAME")
            ]
            track_input = None
            used = ""
            for by, sel, desc in track_selectors:
                el = check_element(driver, by, sel, 3, desc)
                if el:
                    track_input = el; used = desc; break
            if not track_input:
                log_error("Tracking mező nem talalhato"); continue
            track_input.clear(); track_input.send_keys(tracking)
            log_success(f"Tracking szám beirva ({used})")
            time.sleep(1)

            log_step("3b", "Track gomb keresése...")
            btn_selectors = [
                (By.ID, "stApp_btnTrack", "ID: stApp_btnTrack"),
                (By.XPATH, "//button[contains(text(),'Track')]", "Szöveg: Track"),
                (By.CSS_SELECTOR, "button[type='submit']", "Type submit")
            ]
            track_btn = None
            for by, sel, desc in btn_selectors:
                el = check_element(driver, by, sel, 3, desc)
                if el and el.is_enabled():
                    track_btn = el; used = desc; break
            if not track_btn:
                log_error("Track gomb nem talalhato"); continue
            track_btn.click()
            log_success(f"Track gomb megnyomva ({used})")

            close_chat_if_present(driver)

            log_step("3c", "Proof of Delivery link keresese...")
            pod_selectors = [
                (By.ID, "stApp_btnProofOfDeliveryonDetails", "ID: stApp_btnProofOfDeliveryonDetails"),
                (By.LINK_TEXT, "Proof of Delivery", "Link szöveg"),
                (By.PARTIAL_LINK_TEXT, "Proof", "Reszleges")
            ]
            pod_link = None
            for by, sel, desc in pod_selectors:
                el = check_element(driver, by, sel, 10, desc)
                if el:
                    pod_link = el; used = desc; break
            if not pod_link:
                log_error("POD link nem talalhato"); continue

            main_window = driver.current_window_handle
            pod_link.click()
            log_success(f"POD link megnyitva ({used})")

            log_step("3d", "Ablakvaltas...")
            try:
                WebDriverWait(driver, 5).until(lambda d: len(d.window_handles) > 1)
                for w in driver.window_handles:
                    if w != main_window:
                        driver.switch_to.window(w); break
                log_success("Uj ablakra valtva")
                time.sleep(2)
            except:
                log_step("Ablak", "Nincs uj ablak, maradunk")

            log_step("3e", "Print this page kereses...")
            print_selectors = [
                (By.ID, "stApp_POD_btnPrint", "ID: stApp_POD_btnPrint"),
                (By.LINK_TEXT, "Print this page", "Link szöveg")
            ]
            print_link = None
            for by, sel, desc in print_selectors:
                el = check_element(driver, by, sel, 5, desc)
                if el:
                    print_link = el; used = desc; break
            if print_link:
                print_link.click()
                log_success(f"Print link megnyitva ({used})")
                time.sleep(2)

            handle_chrome_print(driver)

            driver.switch_to.window(main_window)

            log_step("3f", "Letoltott fajl kereses...")
            time.sleep(3)
            files = os.listdir(download_folder)
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            if pdfs:
                full_paths = [os.path.join(download_folder, f) for f in pdfs]
                latest = max(full_paths, key=os.path.getctime)
                new_path = os.path.join(download_folder, f"{new_name}.pdf")
                if os.path.exists(new_path): os.remove(new_path)
                shutil.move(latest, new_path)
                log_success(f"Fajl mentve: {new_name}.pdf")
                
                for col in range(1, 6):
                    ws.cell(row=excel_row, column=col).fill = zold_fill
                log_success(f"Sor {excel_row} zoldre szinezve (A-E oszlopok, #{GREEN_COLOR})")
                
                success_count += 1
            else:
                log_error("Nem talalhato letoltott PDF")

            processed += 1
            update_progress(processed, total)
            log_success(f"Feldolgozva: {processed}/{total}")

        log_message("\n[4/5] Excel fajl mentese...")
        output_path = excel_path.replace('.xlsx', '_FELDOLGOZOTT.xlsx')
        if output_path == excel_path:
            output_path = excel_path + '_FELDOLGOZOTT.xlsx'
        
        try:
            wb.save(output_path)
            log_success(f"Excel mentve: {output_path}")
            log_message(f"Sikeres: {success_count}/{total}\n")
        except Exception as e:
            log_error("Excel mentesi hiba", str(e))
            return 1

        log_message("[5/5] Folyamat befejezve")
        return 0
    except Exception as e:
        log_error("Varatlan hiba", str(e)); return 1
    finally:
        if driver:
            driver.quit()
            log_message("Bongeszo bezarva")
        if os.path.exists(STOP_FILE): os.remove(STOP_FILE)

if __name__ == "__main__":
    sys.exit(main())
'@
    
    $tempPython = [System.IO.Path]::GetTempFileName() + ".py"
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tempPython, $pythonScript, $utf8WithBom)
    
    Write-Log "Python script futtatasa..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "python"
    $psi.Arguments = "`"$tempPython`" `"$url`" `"$excelPath`" `"$downloadFolder`""
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
    $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
    
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
    
    Write-Log ""; Write-Log "="*50
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

# Kilépés gomb
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(520, 580)
$exitButton.Size = New-Object System.Drawing.Size(90, 25)
$exitButton.Text = "Kilépés"
$exitButton.BackColor = "DarkRed"
$exitButton.ForeColor = "White"
$exitButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$exitButton.Add_Click({
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
        Set-Content -Path $stopFilePath -Value "stop" -Force
        Write-Log "Leallitasi jelzes kuldve..."
        Start-Sleep -Seconds 2
        if (!$script:pythonProcess.HasExited) { $script:pythonProcess.Kill() }
    }
    $form.Close()
})
$form.Controls.Add($exitButton)

$form.ShowDialog()