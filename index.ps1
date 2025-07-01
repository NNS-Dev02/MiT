Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$folder = "C:\MiT"
if (-not (Test-Path $folder)) { New-Item -Path $folder -ItemType Directory | Out-Null }

$imgPath  = Join-Path $folder "logo.png"
$appsPath = Join-Path $folder "apps.txt"

$logoUrl  = "https://raw.githubusercontent.com/NNS-Dev02/MiT/main/logo.png"
$appsUrl  = "https://raw.githubusercontent.com/NNS-Dev02/MiT/main/apps.txt"

if (-not (Test-Path $imgPath)) { try { Invoke-WebRequest $logoUrl -OutFile $imgPath -UseBasicParsing } catch { } }
if (-not (Test-Path $appsPath)) {
    try { Invoke-WebRequest $appsUrl -OutFile $appsPath -UseBasicParsing }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Không tải được file apps.txt.", "Lỗi", "OK", "Error")
        exit
    }
}

$appListRaw = Get-Content $appsPath -Raw
$appList = $appListRaw -split "`n" | Where-Object { $_ -match "^\s*(.+?)\s*\|\s*(.+?)\s*$" }
$menuItems = foreach ($line in $appList) {
    if ($line -match "^\s*(.+?)\s*\|\s*(.+?)\s*$") {
        [PSCustomObject]@{ Name = $matches[1].Trim(); Value = $matches[2].Trim() }
    }
}

$form = New-Object Windows.Forms.Form
$form.Text = "Công Ty TNHH Thương Mại Và Sản Xuất MiT"
$form.Size = New-Object Drawing.Size(600, 750)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object Drawing.Font("Segoe UI", 12)

$pictureBox = New-Object Windows.Forms.PictureBox
$pictureBox.Size = New-Object Drawing.Size(200, 150)
$pictureBox.SizeMode = "StretchImage"
$pictureBox.Location = New-Object Drawing.Point(([int](($form.ClientSize.Width - 200) / 2)), 10)
if (Test-Path $imgPath) { try { $pictureBox.Image = [System.Drawing.Image]::FromFile($imgPath) } catch { } }
$form.Controls.Add($pictureBox)

$listBox = New-Object Windows.Forms.CheckedListBox
$listBox.Size = New-Object Drawing.Size(560, 450)
$listBox.Location = New-Object Drawing.Point(10, 180)
$listBox.CheckOnClick = $true
$listBox.Font = New-Object Drawing.Font("Segoe UI", 12)
$menuItems.ForEach({ $listBox.Items.Add($_.Name) })
$form.Controls.Add($listBox)

$global:doInstall = $false

$okButton = New-Object Windows.Forms.Button
$okButton.Text = "Cài đặt"
$okButton.Location = New-Object Drawing.Point(300, 650)
$okButton.Size = New-Object Drawing.Size(80, 30)
$okButton.Font = New-Object Drawing.Font("Segoe UI", 11)
$okButton.Add_Click({
    $global:doInstall = $true
    $form.Close()
})
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button
$cancelButton.Text = "Thoát"
$cancelButton.Location = New-Object Drawing.Point(390, 650)
$cancelButton.Size = New-Object Drawing.Size(80, 30)
$cancelButton.Font = New-Object Drawing.Font("Segoe UI", 11)
$cancelButton.Add_Click({
    $global:doInstall = $false
    $form.Close()
})
$form.Controls.Add($cancelButton)

$form.Topmost = $true
$form.ShowDialog()

if (-not $global:doInstall) {
    Write-Host "Bạn đã chọn Thoát. Không thực hiện gì cả." -ForegroundColor Cyan
    exit
}

$selectedItems = $listBox.CheckedItems
if (-not $selectedItems -or $selectedItems.Count -eq 0) {
    Write-Host "Không có lựa chọn nào được chọn. Thoát." -ForegroundColor Yellow
    exit
}

# Hàm kiểm tra và cài winget nếu chưa có
function Install-WingetIfNeeded {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "Winget chưa có trên máy, sẽ tiến hành cài đặt..." -ForegroundColor Yellow

        $wingetInstallerUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle"
        $tempFolder = "$env:TEMP\winget_install"
        if (-not (Test-Path $tempFolder)) { New-Item $tempFolder -ItemType Directory | Out-Null }
        $appInstallerPath = Join-Path $tempFolder "Microsoft.DesktopAppInstaller.appxbundle"

        Write-Host "Đang tải App Installer..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $wingetInstallerUrl -OutFile $appInstallerPath -UseBasicParsing

        Write-Host "Đang cài đặt winget..." -ForegroundColor Cyan
        try {
            Add-AppxPackage -Path $appInstallerPath
            Write-Host "Cài đặt winget thành công. Bạn có thể cần khởi động lại PowerShell." -ForegroundColor Green
        } catch {
            Write-Error "Lỗi khi cài winget: $_"
            Write-Host "Bạn hãy cài winget thủ công theo https://aka.ms/getwinget"
            exit
        }
    } else {
        Write-Host "Winget đã được cài đặt." -ForegroundColor Green
    }
}

Install-WingetIfNeeded

foreach ($selName in $selectedItems) {
    $entry = $menuItems | Where-Object { $_.Name -eq $selName }
    $val = $entry.Value

    switch ($val) {
        "script" {
            Write-Host "`nĐang chạy script MAS..." -ForegroundColor Yellow
            irm https://get.activated.win | iex
        }
        "link" {
            Write-Host "`nĐang mở link tải Office..." -ForegroundColor Cyan
            Start-Process "https://drive.google.com/file/d/1dEfb8xFzNeLMhm5ji9W9n07zyc1WXRRG/view?usp=sharing"
        }
        default {
            if (Get-Command winget -ErrorAction SilentlyContinue) {
                Write-Host "`nĐang cài đặt: $val" -ForegroundColor Green
                winget install --id $val -e --accept-package-agreements --accept-source-agreements
            } else {
                Write-Warning "Không tìm thấy lệnh winget. Vui lòng cài đặt Windows Package Manager."
            }
        }
    }
}