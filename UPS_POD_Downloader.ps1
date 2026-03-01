# UPS_POD_Downloader.ps1
# UPS Proof of Delivery automatizált letöltő
# Futtatás: Jobb klikk -> Run with PowerShell

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# GUI létrehozása
$form = New-Object System.Windows.Forms.Form
$form.Text = "UPS POD Letöltő"
$form.Size = New-Object System.Drawing.Size(650, 700)  # Kicsit nagyobb a stop gomb miatt
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
    "✓ 'összefüz' - a letöltött fájl végső neve",
    "✓ 'POD feltöltve' - ha üres, feldolgozzuk; ha 'OK', kihagyjuk"
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
$progressBar.Size = New-Object System.Drawing.Size(280, 25)  # Kisebb, hogy elférjen mellette a stop gomb
$form.Controls.Add($progressBar)

# ============================================
# STOP GOMB - ÚJ!
# ============================================
$script:stopRequested = $false
$script:pythonProcess = $null

$stopButton = New-Object System.Windows.Forms.Button
$stopButton.Location = New-Object System.Drawing.Point(300, 540)  # Progress bar mellett
$stopButton.Size = New-Object System.Drawing.Size(90, 25)
$stopButton.Text = "🛑 Megállítás"
$stopButton.BackColor = "Orange"
$stopButton.ForeColor = "White"
$stopButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$stopButton.Enabled = $false  # Kezdetben tiltva, csak futás közben aktív
$stopButton.Add_Click({
    $script:stopRequested = $true
    Write-Log "⚠️ Leállítás kérve... (a következő tracking után leáll)"
    
    # Ha van futó Python folyamat, jelzőfájlt hozunk létre
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

# Függvény a naplózáshoz
function Write-Log {
    param($Message)
    $logBox.AppendText($Message + "`r`n")
    $logBox.Refresh()
    Start-Sleep -Milliseconds 10
}

# Indítás gomb eseménykezelő
$startButton.Add_Click({
    $startButton.Enabled = $false
    $stopButton.Enabled = $true  # Stop gomb engedélyezése
    $script:stopRequested = $false
    
    # Régi stop jelzőfájl törlése, ha van
    $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
    if (Test-Path $stopFilePath) {
        Remove-Item $stopFilePath -Force
    }
    
    $url = $urlBox.Text.Trim()
    $excelPath = $excelBox.Text.Trim()
    $downloadFolder = $folderBox.Text.Trim()
    
    # Ellenőrzések
    if (-not $url) {
        [System.Windows.Forms.MessageBox]::Show("Add meg az UPS URL-t!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        return
    }
    
    if (-not $excelPath -or -not (Test-Path $excelPath)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes Excel fájlt kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        return
    }
    
    if (-not $downloadFolder -or -not (Test-Path $downloadFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Érvényes letöltési mappát kell kiválasztani!", "Hiba", "OK", "Error")
        $startButton.Enabled = $true
        $stopButton.Enabled = $false
        return
    }
    
    Write-Log "==========================================="
    Write-Log "🚀 UPS POD Letöltő indítása"
    Write-Log "==========================================="
    Write-Log "Dátum: $(Get-Date)"
    Write-Log "Excel: $excelPath"
    Write-Log "Letöltési mappa: $downloadFolder"
    Write-Log "UPS URL: $url"
    Write-Log ""
    
    # Python script létrehozása ideiglenes fájlban - FRISSÍTETT VERZIÓ STOP TÁMOGATÁSSAL
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
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# STOP fájl ellenőrzéséhez
STOP_FILE = os.path.join(os.environ['TEMP'], 'ups_pod_stop.txt')

def should_stop():
    """Ellenőrzi, hogy kérték-e a leállítást"""
    return os.path.exists(STOP_FILE)

def log_message(msg):
    """Üzenet küldése a PowerShell-nek"""
    print(f"LOG: {msg}")
    sys.stdout.flush()

def log_error(msg, details=""):
    """Hibaüzenet küldése"""
    print(f"LOG: ❌ {msg}")
    if details:
        print(f"LOG:   🔍 {details}")
    sys.stdout.flush()

def log_success(msg):
    """Sikeres művelet jelzése"""
    print(f"LOG: ✅ {msg}")
    sys.stdout.flush()

def log_step(step, msg):
    """Lépés jelzése"""
    print(f"LOG:   📍 [{step}] {msg}")
    sys.stdout.flush()

def update_progress(current, total):
    """Progress frissítése"""
    print(f"PROGRESS: {current},{total}")
    sys.stdout.flush()

def check_element(driver, by, selector, timeout=5, description=""):
    """Elem keresése részletes hibanaplózással"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
        log_step("Keresés", f"✅ Megtalálva: {description} ({selector})")
        return element
    except TimeoutException:
        log_error(f"Nem található: {description}", f"Selector: {selector}, időtúllépés: {timeout}s")
        return None
    except Exception as e:
        log_error(f"Hiba a kereséskor: {description}", str(e))
        return None

def main():
    # Argumentumok beolvasása
    if len(sys.argv) < 4:
        log_error("Hiányzó argumentumok")
        return 1
    
    ups_url = sys.argv[1]
    excel_path = sys.argv[2]
    download_folder = sys.argv[3]
    
    log_message("=" * 60)
    log_message("🔧 PYTHON SCRIPT FUT")
    log_message("=" * 60)
    log_message(f"📂 Excel: {excel_path}")
    log_message(f"📁 Mappa: {download_folder}")
    log_message(f"🌐 URL: {ups_url}")
    log_message("")
    
    # =========================================
    # 1. LÉPÉS: Excel beolvasása
    # =========================================
    log_message("📊 [1/5] Excel fájl beolvasása...")
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        log_success(f"Excel beolvasva - {len(df)} sor, oszlopok: {list(df.columns)}")
    except FileNotFoundError:
        log_error("Excel fájl nem található", f"Útvonal: {excel_path}")
        return 1
    except Exception as e:
        log_error("Excel olvasási hiba", str(e))
        return 1
    
    # Szükséges oszlopok ellenőrzése
    required = ['Tracking Number', 'összefüz', 'POD feltöltve']
    missing = [col for col in required if col not in df.columns]
    if missing:
        log_error("Hiányzó oszlopok", f"Kell: {required}, Hiányzik: {missing}")
        return 1
    
    # Feldolgozandó sorok szűrése
    to_process = df[df['POD feltöltve'].isna() | (df['POD feltöltve'] == '')]
    total = len(to_process)
    
    if total == 0:
        log_message("ℹ️ Nincs feldolgozandó sor.")
        return 0
    
    log_success(f"Feldolgozandó sorok: {total}")
    update_progress(0, total)
    log_message("")
    
    # =========================================
    # 2. LÉPÉS: Böngésző indítása
    # =========================================
    log_message("🌐 [2/5] Böngésző indítása...")
    chrome_options = Options()
    
    # Letöltési beállítások
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
    
    # Chrome profil használata
    user_data_dir = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Google', 'Chrome', 'User Data')
    if os.path.exists(user_data_dir):
        chrome_options.add_argument(f"--user-data-dir={user_data_dir}")
        chrome_options.add_argument("--profile-directory=Default")
        log_step("Profil", "Meglévő Chrome profil betöltve")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        log_success("Böngésző sikeresen elindult")
    except WebDriverException as e:
        log_error("Böngésző indítási hiba", 
                  "Ellenőrizd: Chrome telepítve van? A driver verzió megfelelő?")
        return 1
    
    try:
        # UPS oldal megnyitása
        log_step("Oldal", f"Megnyitás: {ups_url}")
        driver.get(ups_url)
        time.sleep(3)
        log_success("Oldal betöltve")
        log_message("")
        
        processed = 0
        success_count = 0
        
        # =========================================
        # 3. LÉPÉS: Tracking számok feldolgozása
        # =========================================
        for idx, row in to_process.iterrows():
            # Ellenőrizzük, hogy kérték-e a leállítást
            if should_stop():
                log_message("⚠️ Leállítási kérés észlelve (STOP fájl)")
                log_message("   A folyamat megszakítása a felhasználó kérésére...")
                # STOP fájl törlése
                if os.path.exists(STOP_FILE):
                    os.remove(STOP_FILE)
                break
            
            tracking = str(row['Tracking Number']).strip()
            new_name = str(row['összefüz']).strip()
            
            log_message("")
            log_message("─" * 50)
            log_message(f"📦 Feldolgozás: {tracking} -> {new_name}")
            log_message("─" * 50)
            
            # =========================================
            # 3a. Tracking szám mező keresése és kitöltése
            # =========================================
            log_step("3a", "Tracking szám mező keresése...")
            
            # TÖBBSZÖRÖS SELECTOR PRÓBÁLKOZÁS
            tracking_selectors = [
                (By.ID, "st_app_trackingnumber", "ID: st_app_trackingnumber"),
                (By.ID, "stApp_trackingNumber", "ID: stApp_trackingNumber"),
                (By.NAME, "trackingnumber", "NAME: trackingnumber"),
                (By.CSS_SELECTOR, "textarea[formcontrolname='trackingNumber']", "Angular form control"),
                (By.CSS_SELECTOR, "textarea.ups-textbox_textarea", "Class: ups-textbox_textarea"),
                (By.CSS_SELECTOR, "input[placeholder*='Tracking']", "Placeholder: Tracking"),
                (By.CSS_SELECTOR, "[aria-label*='track']", "ARIA label"),
                (By.XPATH, "//textarea[contains(@id, 'tracking')]", "XPATH: ID tartalmaz tracking")
            ]
            
            track_input = None
            selector_used = ""
            for by, selector, desc in tracking_selectors:
                element = check_element(driver, by, selector, timeout=3, description=desc)
                if element:
                    track_input = element
                    selector_used = desc
                    break
            
            if not track_input:
                log_error("Tracking szám mező nem található", 
                         "Az UPS oldala változott - ellenőrizd a selectorokat")
                # DEBUG: Oldal forrásának mentése
                with open("debug_page.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
                log_step("Debug", "Oldal forrása elmentve: debug_page.html")
                continue
            
            # Tracking szám beírása
            try:
                track_input.clear()
                track_input.send_keys(tracking)
                log_success(f"Tracking szám beírva (selector: {selector_used})")
                time.sleep(1)
            except Exception as e:
                log_error("Nem sikerült beírni a tracking számot", str(e))
                continue
            
            # =========================================
            # 3b. Track gomb keresése és kattintás
            # =========================================
            log_step("3b", "Track gomb keresése...")
            
            track_button_selectors = [
                (By.ID, "st_app_track_package_btn", "ID: st_app_track_package_btn"),
                (By.ID, "stApp_trackButton", "ID: stApp_trackButton"),
                (By.XPATH, "//button[contains(text(),'Track')]", "Szöveg: Track"),
                (By.XPATH, "//button[contains(text(),'Követés')]", "Szöveg: Követés"),
                (By.CSS_SELECTOR, "button[type='submit']", "Type: submit"),
                (By.CSS_SELECTOR, "button.ups-button_primary", "Class: ups-button_primary")
            ]
            
            track_btn = None
            selector_used = ""
            for by, selector, desc in track_button_selectors:
                element = check_element(driver, by, selector, timeout=3, description=desc)
                if element and element.is_enabled():
                    track_btn = element
                    selector_used = desc
                    break
            
            if not track_btn:
                log_error("Track gomb nem található", "Az UPS oldala változott")
                continue
            
            try:
                track_btn.click()
                log_success(f"Track gomb megnyomva (selector: {selector_used})")
                time.sleep(4)
            except Exception as e:
                log_error("Nem sikerült a Track gombra kattintani", str(e))
                continue
            
            # =========================================
            # 3c. Proof of Delivery link keresése
            # =========================================
            log_step("3c", "Proof of Delivery link keresése...")
            
            pod_selectors = [
                (By.LINK_TEXT, "Proof of Delivery", "Link szöveg: Proof of Delivery"),
                (By.PARTIAL_LINK_TEXT, "Proof", "Részleges szöveg: Proof"),
                (By.XPATH, "//a[contains(text(),'Proof')]", "XPATH: szöveg tartalmaz Proof"),
                (By.CSS_SELECTOR, "a[href*='proof']", "HREF tartalmaz proof"),
                (By.XPATH, "//a[contains(@class, 'pod-link')]", "Class: pod-link")
            ]
            
            pod_link = None
            selector_used = ""
            for by, selector, desc in pod_selectors:
                element = check_element(driver, by, selector, timeout=5, description=desc)
                if element:
                    pod_link = element
                    selector_used = desc
                    break
            
            if not pod_link:
                log_error("Proof of Delivery link nem található", 
                         "Lehet, hogy nincs POD ehhez a csomaghoz")
                continue
            
            # Főablak azonosítója
            main_window = driver.current_window_handle
            
            try:
                pod_link.click()
                log_success(f"POD link megnyitva (selector: {selector_used})")
                time.sleep(2)
            except Exception as e:
                log_error("Nem sikerült a POD linkre kattintani", str(e))
                continue
            
            # =========================================
            # 3d. Új ablakra váltás
            # =========================================
            log_step("3d", "Ablakváltás...")
            try:
                all_windows = driver.window_handles
                if len(all_windows) > 1:
                    for window in all_windows:
                        if window != main_window:
                            driver.switch_to.window(window)
                            break
                    log_success(f"Új ablakra váltva ({len(all_windows)} ablak nyitva)")
                else:
                    log_step("Ablak", "Nincs új ablak, maradunk a főablakban")
                time.sleep(2)
            except Exception as e:
                log_error("Ablakváltási hiba", str(e))
            
            # =========================================
            # 3e. Print this page link keresése
            # =========================================
            log_step("3e", "Print this page link keresése...")
            
            print_selectors = [
                (By.LINK_TEXT, "Print this page", "Link szöveg: Print this page"),
                (By.PARTIAL_LINK_TEXT, "Print", "Részleges szöveg: Print"),
                (By.XPATH, "//a[contains(text(),'Print')]", "XPATH: szöveg tartalmaz Print"),
                (By.ID, "printLink", "ID: printLink"),
                (By.CSS_SELECTOR, "button[onclick*='print']", "Gomb onclick tartalmaz print")
            ]
            
            print_link = None
            selector_used = ""
            for by, selector, desc in print_selectors:
                element = check_element(driver, by, selector, timeout=5, description=desc)
                if element:
                    print_link = element
                    selector_used = desc
                    break
            
            if print_link:
                try:
                    print_link.click()
                    log_success(f"Print link megnyitva (selector: {selector_used})")
                    time.sleep(2)
                except Exception as e:
                    log_error("Nem sikerült a Print linkre kattintani", str(e))
            else:
                log_step("Print", "Nincs Print link, lehet hogy közvetlen PDF")
            
            # =========================================
            # 3f. Save gomb keresése
            # =========================================
            log_step("3f", "Save gomb keresése...")
            
            save_selectors = [
                (By.ID, "save", "ID: save"),
                (By.XPATH, "//button[contains(text(),'Save')]", "Gomb szöveg: Save"),
                (By.XPATH, "//input[@value='Save']", "Input value: Save"),
                (By.CSS_SELECTOR, "button[onclick*='save']", "Onclick tartalmaz save"),
                (By.XPATH, "//button[contains(@class, 'save')]", "Class tartalmaz save")
            ]
            
            save_btn = None
            selector_used = ""
            for by, selector, desc in save_selectors:
                element = check_element(driver, by, selector, timeout=3, description=desc)
                if element and element.is_enabled():
                    save_btn = element
                    selector_used = desc
                    break
            
            if save_btn:
                try:
                    save_btn.click()
                    log_success(f"Save gomb megnyomva (selector: {selector_used})")
                    time.sleep(3)
                except Exception as e:
                    log_error("Nem sikerült a Save gombra kattintani", str(e))
            else:
                log_step("Save", "Nincs Save gomb, lehet hogy automatikus a letöltés")
            
            # Visszaváltás a főablakra
            try:
                driver.switch_to.window(main_window)
                log_step("Ablak", "Visszaváltva a főablakra")
            except:
                pass
            
            # =========================================
            # 3g. Letöltött fájl keresése és átnevezése
            # =========================================
            log_step("3g", "Letöltött fájl keresése...")
            time.sleep(2)
            
            try:
                files = os.listdir(download_folder)
                pdf_files = [f for f in files if f.lower().endswith('.pdf')]
                
                if pdf_files:
                    # Legutóbb módosított PDF kiválasztása
                    pdf_files_with_path = [os.path.join(download_folder, f) for f in pdf_files]
                    latest_pdf = max(pdf_files_with_path, key=os.path.getctime)
                    
                    log_step("Fájl", f"Legújabb PDF: {os.path.basename(latest_pdf)}")
                    
                    # Átnevezés
                    new_path = os.path.join(download_folder, f"{new_name}.pdf")
                    if os.path.exists(new_path):
                        os.remove(new_path)
                        log_step("Fájl", "Régi fájl törölve")
                    
                    shutil.move(latest_pdf, new_path)
                    log_success(f"Fájl mentve: {new_name}.pdf")
                    
                    # Excel frissítése
                    df.loc[idx, 'POD feltöltve'] = 'OK'
                    success_count += 1
                else:
                    log_error("Nem található letöltött PDF", 
                             "Ellenőrizd a letöltési mappát és a Chrome beállításokat")
                
            except Exception as e:
                log_error("Hiba a fájl kezelésekor", str(e))
            
            processed += 1
            update_progress(processed, total)
            log_success(f"Feldolgozva: {processed}/{total}")
        
        # =========================================
        # 4. LÉPÉS: Excel mentése
        # =========================================
        log_message("")
        log_message("💾 [4/5] Excel fájl mentése...")
        
        try:
            # Eredeti fájl mentése
            output_path = excel_path.replace('.xlsx', '_FELDOLGOZOTT.xlsx')
            if output_path == excel_path:
                output_path = excel_path + '_FELDOLGOZOTT.xlsx'
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                
                # Zöld háttér az OK soroknak
                from openpyxl.styles import PatternFill
                green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
                
                worksheet = writer.sheets['Sheet1']
                for idx, row in df.iterrows():
                    if row['POD feltöltve'] == 'OK':
                        excel_row = idx + 2
                        for col in range(1, len(df.columns) + 1):
                            cell = worksheet.cell(row=excel_row, column=col)
                            cell.fill = green_fill
            
            log_success(f"Excel mentve: {output_path}")
            log_message(f"📊 Sikeres letöltések: {success_count}/{total}")
            
        except Exception as e:
            log_error("Excel mentési hiba", str(e))
            return 1
        
        # =========================================
        # 5. LÉPÉS: Befejezés
        # =========================================
        log_message("")
        log_message("✅ [5/5] Folyamat befejezve")
        return 0
        
    except Exception as e:
        log_error("Váratlan hiba", str(e))
        return 1
    finally:
        if driver:
            driver.quit()
            log_message("🟢 Böngésző bezárva")
        # STOP fájl törlése, ha maradt
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)

if __name__ == "__main__":
    sys.exit(main())
'@
    
    # Ideiglenes Python fájl létrehozása
    $tempPython = [System.IO.Path]::GetTempFileName() + ".py"
    
    # UTF-8-BOM mentés
    $utf8WithBom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($tempPython, $pythonScript, $utf8WithBom)
    
    Write-Log "🚀 Python script futtatása..."
    Write-Log ""
    
    # Python futtatása és kimenet olvasása valós időben
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
    $script:pythonProcess = $process  # Eltároljuk a folyamatot
    
    # Eseménykezelők a kimenet olvasásához
    $outputEvent = Register-ObjectEvent -InputObject $process -EventName 'OutputDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            if ($data.StartsWith("LOG: ")) {
                $message = $data.Substring(5)
                $form.BeginInvoke([Action]{ Write-Log $message })
            }
            elseif ($data.StartsWith("PROGRESS: ")) {
                $parts = $data.Substring(10).Split(',')
                if ($parts.Count -eq 2) {
                    $current = [int]$parts[0]
                    $total = [int]$parts[1]
                    $form.BeginInvoke([Action]{ 
                        $progressBar.Maximum = $total
                        $progressBar.Value = $current
                    })
                }
            }
        }
    }
    
    $errorEvent = Register-ObjectEvent -InputObject $process -EventName 'ErrorDataReceived' -Action {
        $data = $EventArgs.Data
        if ($data -ne $null) {
            $form.BeginInvoke([Action]{ Write-Log "❌ PYTHON HIBA: $data" })
        }
    }
    
    # Folyamat indítása
    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    
    # Várakozás a befejeződésre
    $process.WaitForExit()
    $exitCode = $process.ExitCode
    $script:pythonProcess = $null  # Folyamat eltávolítása
    
    # Eseménykezelők eltávolítása
    Unregister-Event -SourceIdentifier $outputEvent.Name -Force -ErrorAction SilentlyContinue
    Unregister-Event -SourceIdentifier $errorEvent.Name -Force -ErrorAction SilentlyContinue
    
    # Ideiglenes fájl törlése
    Remove-Item $tempPython -Force -ErrorAction SilentlyContinue
    
    Write-Log ""
    Write-Log "=" * 50
    if ($exitCode -eq 0) {
        Write-Log "✅ FOLYAMAT SIKERESEN BEFEJEZŐDÖTT"
        [System.Windows.Forms.MessageBox]::Show("A letöltés sikeresen befejeződött!", "Siker", "OK", "Information")
    } else {
        Write-Log "❌ HIBA TÖRTÉNT (kód: $exitCode)"
        Write-Log "   Nézd át a naplót a részletekért!"
        [System.Windows.Forms.MessageBox]::Show("Hiba történt a letöltés során! Ellenőrizd a naplót.", "Hiba", "OK", "Error")
    }
    Write-Log "=" * 50
    
    $progressBar.Value = 0
    $startButton.Enabled = $true
    $stopButton.Enabled = $false  # Stop gomb letiltása
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
    # Ha van futó Python folyamat, próbáljuk meg leállítani
    if ($script:pythonProcess -and !$script:pythonProcess.HasExited) {
        $stopFilePath = Join-Path $env:TEMP "ups_pod_stop.txt"
        Set-Content -Path $stopFilePath -Value "stop" -Force
        Write-Log "⚠️ Leállítási jelzés küldve a Python folyamatnak..."
        Start-Sleep -Seconds 2
        if (!$script:pythonProcess.HasExited) {
            $script:pythonProcess.Kill()
            Write-Log "🛑 Python folyamat kényszerített leállítása"
        }
    }
    $form.Close() 
})
$form.Controls.Add($exitButton)

# Form megjelenítése
$form.ShowDialog()