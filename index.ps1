# Bật Visual Styles cho WinForms
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()

# Khai báo API để dùng SetForegroundWindow
if (-not ([System.Type]::GetType("Native.WinApi"))) {
    Add-Type -MemberDefinition @"
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
"@ -Name WinApi -Namespace Native
}

# Yêu cầu nhập mật khẩu
$securePassword = Read-Host "Nhập mật khẩu để tiếp tục" -AsSecureString
$plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
)

# Mật khẩu đúng (ví dụ: "mit2025")
$correctPassword = "mitgroup2025"

if ($plainPassword -ne $correctPassword) {
    Write-Host "Sai mật khẩu. Không thể tiếp tục." -ForegroundColor Red
    exit
}





#---------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------Bảng thông báo--------------------------------------------------
#---------------------------------------------------------------------------------------------------------------------
# 1. Hàm hiển thị dialog "Đang tải..."
function Show-LoadingDialog {
    param([string] $Message = "Đang tải...")
    $dlg = New-Object Windows.Forms.Form
    $dlg.Text              = "Thông báo"
    $dlg.Size              = New-Object Drawing.Size(300,100)
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.StartPosition   = "CenterScreen"
    $dlg.Topmost           = $true
    $dlg.ControlBox        = $false

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text     = $Message
    $lbl.AutoSize = $true
    $lbl.Font     = New-Object Drawing.Font("Segoe UI",10)
    $lbl.Location = New-Object Drawing.Point(
        [int](($dlg.ClientSize.Width - $lbl.PreferredWidth)/2),
        [int](($dlg.ClientSize.Height - $lbl.PreferredHeight)/2)
    )
    $dlg.Controls.Add($lbl)
    $dlg.Show()
    return $dlg
}
# 2. Hàm tải file với loading dialog
function Download-WithLoading {
    param([string] $Url, [string] $OutFile, [string] $Message)
    $dlg = Show-LoadingDialog $Message
    try { Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing }
    catch {
        [Windows.Forms.MessageBox]::Show("Lỗi khi tải file: $($_.Exception.Message)", "Lỗi", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
    }
    $dlg.Close()
}





#---------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------Chức năng bảo trì--------------------------------------------------
#--------------------------------------------------Hàm xuất file Log--------------------------------------------------
# Hàm hiển thị thông báo đếm ngược
function Show-CountdownMessageDialog {
    $msgForm = New-Object System.Windows.Forms.Form
    $msgForm.Text              = "Thông báo"
    $msgForm.Size              = New-Object System.Drawing.Size(450, 180)
    $msgForm.StartPosition   = "CenterScreen"
    $msgForm.FormBorderStyle = "FixedDialog"
    $msgForm.MinimizeBox     = $false
    $msgForm.MaximizeBox     = $false
    $msgForm.ControlBox        = $false
    $msgForm.Topmost           = $true

    $mainLabel = New-Object System.Windows.Forms.Label
    $mainLabel.Text     = "Vui lòng không sử dụng bàn phím và chuột đến khi chức năng chạy xong!!!"
    $mainLabel.Location = New-Object System.Drawing.Point(20, 30)
    $mainLabel.Size     = New-Object System.Drawing.Size(400, 60)
    $mainLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $mainLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 11)
    $msgForm.Controls.Add($mainLabel)

    $countdownLabel = New-Object System.Windows.Forms.Label
    $countdownLabel.Text     = "Tự động tắt sau: 5 giây"
    $countdownLabel.Location = New-Object System.Drawing.Point(20, 90)
    $countdownLabel.Size     = New-Object System.Drawing.Size(400, 30)
    $countdownLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $countdownLabel.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $msgForm.Controls.Add($countdownLabel)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Enabled = $true
    
    $timer.Tag = 5 

    $timer.Add_Tick({
        param($sender, $e)
        
        $currentCountdownTime = $sender.Tag
        $currentCountdownTime--
        $sender.Tag = $currentCountdownTime

        Write-Host "Đếm ngược: $currentCountdownTime"
        
        $countdownLabel.Text = "Tự động tắt sau: $currentCountdownTime giây"

        if ($currentCountdownTime -le 0) {
            $sender.Stop()
            $msgForm.Close()
        }
    })

    $msgForm.Add_Shown({
        $timer.Start()
        [Windows.Forms.Application]::DoEvents() 
    })

    $msgForm.ShowDialog()
}
# Hàm Chức năng dịch vụ bảo trì
function Show-SystemInfoForm {

    Show-CountdownMessageDialog

    $dl = Join-Path $env:USERPROFILE "Downloads"

# 1. Chạy ứng dụng Core Temp
$core = "C:\Program Files\Core Temp\Core Temp.exe"
$coreTempInstallDir = Split-Path $core

if (Test-Path $core) {
    Write-Host "----------------------------------------------------------------------------------------------------`n------------------------------Core Temp------------------------------" -ForegroundColor DarkYellow

    $tempCsvInDownloads = Join-Path $dl "temp.csv"
    if (Test-Path $tempCsvInDownloads) {
        Remove-Item $tempCsvInDownloads -ErrorAction SilentlyContinue
        Write-Host "Xóa file temp.csv cũ trong thư mục Downloads." -ForegroundColor DarkYellow
    }

    Get-ChildItem -Path $coreTempInstallDir -Filter "CT-Log*.csv" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -ErrorAction SilentlyContinue
        Write-Host "Xóa file log cũ: $($_.Name)" -ForegroundColor DarkYellow
    }

    $p = Start-Process $core -PassThru
    Start-Sleep 5

    if ($p -and -not $p.HasExited -and $p.MainWindowHandle -ne 0) {
        Write-Host "Phát hiện cửa sổ Core Temp." -ForegroundColor DarkYellow
        [Native.WinApi]::SetForegroundWindow($p.MainWindowHandle)
        Start-Sleep 2

        [System.Windows.Forms.SendKeys]::SendWait("{F4}")
        Write-Host "Gửi phím F4 để bắt đầu ghi log, chờ 10 giây để thu thập dữ liệu..." -ForegroundColor DarkYellow
        Start-Sleep 10

        [System.Windows.Forms.SendKeys]::SendWait("{F4}")
        Write-Host "Gửi phím F4 lần nữa để dừng ghi log." -ForegroundColor DarkYellow
        Start-Sleep 2

        Write-Host "Đóng ứng dụng Core Temp..." -ForegroundColor DarkYellow
        $p.CloseMainWindow()
        Start-Sleep 2
        if (-not $p.HasExited) {
            Stop-Process -Id $p.Id -Force
            Write-Host "Core Temp chưa đóng, đã buộc tắt tiến trình." -ForegroundColor DarkYellow
        }

        $newestCoreTempCsv = $null
        for ($i = 0; $i -lt 5; $i++) {
            $newestCoreTempCsv = Get-ChildItem -Path $coreTempInstallDir -Filter "CT-Log*.csv" | 
                                 Where-Object { $_.LastWriteTime -gt (Get-Date).AddSeconds(-15) } | 
                                 Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($newestCoreTempCsv) { break }
            Start-Sleep 1
        }

        if ($newestCoreTempCsv) {
            Move-Item $newestCoreTempCsv.FullName $tempCsvInDownloads -Force
            Write-Host "Tìm thấy file log của Core Temp: $($newestCoreTempCsv.Name)" -ForegroundColor DarkYellow
            Write-Host "Di chuyển và đổi tên thành temp.csv trong thư mục Downloads." -ForegroundColor DarkYellow
        } else {
            Write-Host "Không tìm thấy file log CSV của Core Temp trong vòng 15 giây gần đây tại thư mục '$coreTempInstallDir'." -ForegroundColor DarkYellow
            Write-Host "Vui lòng kiểm tra thủ công Core Temp để chắc chắn rằng nó đang ghi log bằng phím F4 và xác định vị trí lưu file log." -ForegroundColor DarkYellow
        }

    } else {
        Write-Host "Không phát hiện được cửa sổ Core Temp hoặc chương trình đã đóng sớm." -ForegroundColor DarkYellow
        Write-Host "Vui lòng đảm bảo rằng Core Temp đang chạy và không có cửa sổ/hộp thoại nào chặn lại." -ForegroundColor DarkYellow
    }
} else {
    Write-Host "Không tìm thấy file Core Temp.exe tại đường dẫn: $core" -ForegroundColor DarkYellow
}


# 2. Chạy ứng dụng CPU-Z
$cpuZ = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"

if (Test-Path $cpuZ) {
    Write-Host "----------------------------------------------------------------------------------------------------`n
    ------------------------------CPU-Z------------------------------" -ForegroundColor Magenta

    & $cpuZ -txt=cpu
    Start-Sleep 3

    $cpuLogPath = Join-Path (Split-Path $cpuZ) "cpu.txt"
    $destCpuPath = Join-Path $dl "cpu.txt"

    if (Test-Path $destCpuPath) {
        Remove-Item $destCpuPath -Force -ErrorAction SilentlyContinue
        Write-Host "Đã xóa file cpu.txt cũ trong thư mục Downloads." -ForegroundColor Magenta
    }

    if (Test-Path $cpuLogPath) {
        try {
            Move-Item $cpuLogPath $destCpuPath -Force -ErrorAction Stop
            Write-Host "Di chuyển file cpu.txt vào thư mục Downloads." -ForegroundColor Magenta
        } catch {
            Write-Host "Đã xảy ra lỗi khi di chuyển cpu.txt: $($_.Exception.Message)" -ForegroundColor Magenta
        }

        if (Test-Path $cpuLogPath) {
            Write-Host "File cpu.txt vẫn còn tồn tại tại thư mục gốc sau khi di chuyển (có thể lỗi)." -ForegroundColor Magenta
        } else {
            Write-Host "File cpu.txt đã được di chuyển thành công, không còn ở thư mục gốc." -ForegroundColor Magenta
        }

        if (Test-Path $destCpuPath) {
            Write-Host "File cpu.txt đã tồn tại ở thư mục Downloads." -ForegroundColor Magenta
        } else {
            Write-Host "File cpu.txt KHÔNG tồn tại ở thư mục Downloads sau khi di chuyển." -ForegroundColor Magenta
        }

    } else {
        Write-Host "Không tìm thấy file cpu.txt tại thư mục cài đặt CPU-Z sau khi thực thi." -ForegroundColor Magenta
        Write-Host "Vui lòng kiểm tra thủ công hoặc đảm bảo CPU-Z được chạy với quyền ghi file." -ForegroundColor Magenta
    }

} else {
    Write-Host "Không tìm thấy tệp thực thi CPU-Z tại đường dẫn: $cpuZ" -ForegroundColor Magenta
}


# 3. Chạy ứng dụng CrystalDiskInfo
$cd = "C:\Program Files\CrystalDiskInfo\DiskInfo64.exe"
$diskLogPath = Join-Path (Split-Path $cd) "DiskInfo.txt"
$destDiskPath = Join-Path $dl "disk.txt"

if (Test-Path $cd) {
    Write-Host "----------------------------------------------------------------------------------------------------`n------------------------------CrystalDiskInfo------------------------------" -ForegroundColor Blue

    try {
        Set-Clipboard -Value ""
    } catch {
        Write-Host "Không thể xóa nội dung clipboard: $($_.Exception.Message)" -ForegroundColor Blue
    }

    # Xóa log cũ nếu có
    if (Test-Path $diskLogPath) {
        Remove-Item $diskLogPath -Force -ErrorAction SilentlyContinue
    }

    # Xóa disk.txt trong Downloads nếu có
    if (Test-Path $destDiskPath) {
        Remove-Item $destDiskPath -Force -ErrorAction SilentlyContinue
        Write-Host "Đã xóa file disk.txt cũ trong thư mục Downloads." -ForegroundColor Blue
    }

    # Gọi CrystalDiskInfo với /CopyExit để xuất log và tự thoát
    & $cd /CopyExit

    # Chờ đến khi DiskInfo.txt được tạo (tối đa 10 giây)
    $maxWait = 10
    $waited = 0
    while (-not (Test-Path $diskLogPath) -and $waited -lt $maxWait) {
        Start-Sleep -Seconds 1
        $waited++
    }

    if (Test-Path $diskLogPath) {
        Write-Host "Đã tìm thấy DiskInfo.txt tại thư mục gốc sau $waited giây." -ForegroundColor Blue

        try {
            Move-Item $diskLogPath $destDiskPath -Force -ErrorAction Stop
            Write-Host "Đã di chuyển và đổi tên file DiskInfo.txt thành disk.txt vào thư mục Downloads." -ForegroundColor Blue
        } catch {
            Write-Host "Lỗi khi di chuyển disk.txt: $($_.Exception.Message)" -ForegroundColor Blue
        }

        if (Test-Path $destDiskPath) {
            Write-Host "File disk.txt đã tồn tại tại thư mục Downloads." -ForegroundColor Blue
        } else {
            Write-Host "File disk.txt KHÔNG tồn tại tại thư mục Downloads sau khi di chuyển." -ForegroundColor Blue
        }
    } else {
        Write-Host "Không tìm thấy file DiskInfo.txt sau $maxWait giây. Có thể CrystalDiskInfo chưa kịp ghi log." -ForegroundColor Blue
        Write-Host "Thử tăng thời gian chờ hoặc chạy lại lần nữa nếu cần." -ForegroundColor Blue
    }

} else {
    Write-Host "Không tìm thấy ứng dụng CrystalDiskInfo tại đường dẫn: $cd" -ForegroundColor Blue
}


#------------------------------------------------------------------------------------------------------------------------

#--------------------------------------------------Hàm hiển thị kết quả--------------------------------------------------
# 0. Check Tasks/User :
    $userName = $env:COMPUTERNAME

# 1. Check CPU (Temperature)
function Get-CPUAvgTemp {
    $f = Join-Path $dl "temp.csv"
    if (-not (Test-Path $f)) {
        Write-Host "File temp.csv không tìm thấy tại $f khi cố gắng đọc nhiệt độ." -ForegroundColor DarkYellow
        return "N/A"
    }

    # Write-Host "Đang đọc nội dung file temp.csv (Core Temp log)..." -ForegroundColor DarkYellow
    
    try {
        $fileContentRaw = Get-Content $f -Encoding UTF8 -Raw -ErrorAction Stop
        $fileLines = $fileContentRaw.Split([Environment]::NewLine) | Where-Object { $_.Trim() -ne "" }
    } catch {
        Write-Host "Lỗi khi đọc file temp.csv: $($_.Exception.Message)" -ForegroundColor DarkYellow
        return "N/A"
    }

    if ($fileLines.Count -lt 9) {
        Write-Host "File temp.csv không đủ dòng để đọc nhiệt độ (cần ít nhất 9 dòng)." -ForegroundColor DarkYellow
        Write-Host "Nội dung file:" -ForegroundColor DarkYellow
        $fileLines | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkYellow }
        return "N/A"
    }

    $headerLine = $fileLines[6]
    $dataLines = $fileLines[7..($fileLines.Count - 1)]

    if ($dataLines.Count -gt 0) {
        # Write-Host "Dòng dữ liệu đầu tiên được đọc: '$($dataLines[0])'" -ForegroundColor DarkYellow
    } else {
        Write-Host "Không có dòng dữ liệu nào sau header." -ForegroundColor DarkYellow
        return "N/A"
    }

    $headers = $headerLine -split ','
    $tempColumnIndices = @()

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $trimmedHeader = $headers[$i].Trim()
        if ($trimmedHeader -match '^Core \d+ Temp\.') {
            $tempColumnIndices += $i
        }
    }

    if ($tempColumnIndices.Count -eq 0) {
        Write-Host "Không tìm thấy bất kỳ cột nhiệt độ nào ('Core X Temp.') trong header." -ForegroundColor DarkYellow
        Write-Host "Header đầy đủ: '$headerLine'" -ForegroundColor DarkYellow
        return "N/A"
    }

    $temps = @()
    $firstDataRowValues = ($dataLines[0] -split ',')

    foreach ($index in $tempColumnIndices) {
        if ($index -lt $firstDataRowValues.Count) {
            $val = $firstDataRowValues[$index].Trim()
            if ($val -notmatch '^\s*$' -and $val -notmatch 'N/A' -and $val -notmatch '---') {
                try {
                    $valToConvert = $val -replace '[^0-9\.-]', ''
                    $temps += [double]$valToConvert
                } catch {
                    Write-Host "Không thể chuyển đổi '$val' (từ cột '$($headers[$index])') thành số. Lỗi: $($_.Exception.Message)" -ForegroundColor DarkYellow
                }
            }
        }
    }

    if ($temps.Count -eq 0) {
        Write-Host "Không tìm thấy dữ liệu nhiệt độ hợp lệ từ các cột Core X Temp trong dòng dữ liệu đầu tiên." -ForegroundColor DarkYellow
        return "N/A"
    }

    return ([math]::Round(($temps | Measure-Object -Average).Average, 1)).ToString() + "°C"
}

$cpuTemp = Get-CPUAvgTemp


# 2–3. Check Mainboard : - Check RAM (DDR? & Bus?)
$cpuLog = Join-Path $dl "cpu.txt"
if (Test-Path $cpuLog) {
    $c = Get-Content $cpuLog -Encoding UTF8

    $startIndex = ($c | Select-String '^DMI Baseboard').LineNumber
    if ($startIndex) {
        $block = $c[$startIndex..($startIndex + 10)]

        $vendor = ($block | Where-Object { $_ -match '^\s*vendor\s+' }) -replace '^\s*vendor\s+', ''
        $model  = ($block | Where-Object { $_ -match '^\s*model\s+' })  -replace '^\s*model\s+', ''

        if ($vendor -and $model) {
            $mainboard = "$vendor - $model"
        } else {
            $mainboard = "Không rõ"
        }
    } else {
        $mainboard = "Không rõ"
    }

    # RAM
    $ramSize = ($c | Where-Object { $_ -match '^Memory Size\s+' }) -replace '^.*?\t+', ''
    $ramType = ($c | Where-Object { $_ -match '^Memory Type\s+' }) -replace '^.*?\t+', ''
    $ramBus  = ($c | Where-Object { $_ -match '^Memory Frequency\s+' }) -replace '^.*?\t+', ''
    
    if ($ramBus -match '([\d.]+)\s*MHz') {
        try {
            $actualBusSpeed = [double]$matches[1]
            $effectiveSpeed = $actualBusSpeed * 2
            $ramBus = "Bus $([int][math]::Round($effectiveSpeed))"
        } catch {
            $ramBus = "N/A (Lỗi định dạng bus RAM)"
            Write-Host "Lỗi khi tính toán tốc độ bus hiệu dụng: $($_.Exception.Message)" -ForegroundColor Magenta
        }
    } else {
        $ramBus = "N/A (Không tìm thấy tần số bus RAM)"
    }
} else {
    $mainboard = $ramSize = $ramType = $ramBus = "N/A"
}



# 4. HardDisk Total Space (SSD/HDD):
    $diskLog = Join-Path $dl "disk.txt"
    $diskInfo = ""
    if (Test-Path $diskLog) {
        try {
            $d = Get-Content $diskLog -Encoding UTF8
            $inDiskList = $false
            $lines = @()

            foreach ($line in $d) {
                if ($line -match '^-- Disk List') {
                    $inDiskList = $true
                    continue
                }
                elseif ($inDiskList -and $line -match '^--+') {
                    break
                }
                elseif ($inDiskList -and $line.Trim() -ne "") {
                    if ($line -match '^(.*?)\s*:\s*([\d\.]+\s*GB)') {
                        $prefix = $matches[1].Trim()
                        $size = $matches[2].Trim()
                        $lines += "$prefix : $size"
                    }
                }
            }

            if ($lines.Count -eq 0) {
                $diskInfo = "`r`nKhông có thông tin ổ đĩa"
            } else {
                $diskInfo = "`r`n" + ($lines -join "`r`n")
            }
        } catch {
            $diskInfo = "`r`nLỗi đọc file"
        }
    } else {
        $diskInfo = "`r`nN/A"
    }

# 5. HardDisk Free Space (SSD/HDD):
    $freeStr = Get-CimInstance Win32_LogicalDisk |
        Where-Object { $_.DriveType -eq 3 } |
        ForEach-Object {
            "{0} {1} GB còn trống" -f $_.DeviceID,
                [math]::Round($_.FreeSpace / 1GB, 2)
        }

    $freeStr = ($freeStr | Where-Object { $_.Trim() -ne "" }) -join "`r`n"
    $freeStr = "`r`n" + $freeStr

# 6. Check HardDisk Life :
$diskHealth = ""
if (Test-Path $diskLog) {
    try {
        $d = Get-Content $diskLog -Encoding UTF8
        $healthLines = @()

        foreach ($line in $d) {
            if ($line -match '^\s*Health Status\s*:\s*(.+)$') {
                $healthLines += $matches[1].Trim()
            }
        }

        if ($healthLines.Count -gt 0) {
            $diskHealth = "`r`n" + ($healthLines -join "`r`n")
            # Write-Host "Tình trạng ổ đĩa:" -ForegroundColor Blue
            # Write-Host $diskHealth -ForegroundColor Blue
        } else {
            $diskHealth = "`r`nKhông rõ"
            Write-Host "Tình trạng ổ đĩa: Không rõ" -ForegroundColor Blue
        }
    } catch {
        $diskHealth = "`r`nLỗi đọc file"
        Write-Host "Lỗi khi đọc tình trạng ổ đĩa" -ForegroundColor Blue
    }
} else {
    $diskHealth = "`r`nN/A"
    Write-Host "Không tìm thấy file log tình trạng ổ đĩa" -ForegroundColor Blue
}


# 7. Check Battery Life Laptop :
    $filePath = Join-Path $dl "battery-report.html"
    powercfg /batteryreport /output $filePath > $null
    Start-Sleep -Milliseconds 800
    if (Test-Path $filePath) {
        $html = Get-Content $filePath -Raw

        $designLine = ($html | Select-String -Pattern 'DESIGN CAPACITY.*?(\d[\d,.]*)\s*mWh' -AllMatches).Matches.Value
        $fullLine = ($html | Select-String -Pattern 'FULL CHARGE CAPACITY.*?(\d[\d,.]*)\s*mWh' -AllMatches).Matches.Value

        $dMatch = [regex]::Match($designLine, '\d[\d,.]*')
        $fMatch = [regex]::Match($fullLine, '\d[\d,.]*')

        if ($dMatch.Success -and $fMatch.Success) {
            $d = [double]($dMatch.Value -replace ',','')
            $f = [double]($fMatch.Value -replace ',','')
            $wearPercent = [math]::Round((1 - ($f / $d)) * 100, 1)
            $batteryLife = "Laptop – Hao mòn pin: $wearPercent%"
        } else {
            $batteryLife = "Laptop – Không đọc được dung lượng pin"
        }
    } else {
        $batteryLife = "PC"
    }

# 9. Check Network Connection :
    $ping = Test-Connection -ComputerName 8.8.8.8 -Count 4 -ErrorAction SilentlyContinue
    if ($ping) {
        $avg = ($ping | Measure-Object -Property ResponseTime -Average).Average
        $CheckNetwork = "Độ trễ trung bình : {0:N2} ms" -f $avg
    } else {
        $CheckNetwork = "Không thể ping đến 8.8.8.8"
    }

# 10. Check Date & Time :
    $checkTime = Get-Date -Format "HH:mm:ss dd/MM/yyyy"
# 11. Check Office status :
    $officeStatus = "Không có Office"
    try {
        $word = New-Object -ComObject Word.Application
        $doc = $word.Documents.Add()
        $word.Visible = $false

        if ($word.System.OperatingSystem -ne $null) {
            $officeStatus = "Đã Active"
        } else {
            $officeStatus = "Chưa Active"
        }

        $doc.Close($false)
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    } catch {
        $officeStatus = "Không có Office hoặc không mở được Word"
    }

# 12. Check Mailbox status :
$outlookSizeStatus = ""

# Load WinAPI để điều khiển cửa sổ
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_RESTORE = 9;
}
"@

# Load System.Windows.Forms 1 lần để dùng SendKeys
Add-Type -AssemblyName System.Windows.Forms

function Bring-WindowToFront($processName) {
    $process = Get-Process -Name $processName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($process -and $process.MainWindowHandle -ne 0) {
        [WinAPI]::ShowWindow($process.MainWindowHandle, [WinAPI]::SW_RESTORE) | Out-Null
        [WinAPI]::SetForegroundWindow($process.MainWindowHandle) | Out-Null
        return $true
    }
    return $false
}

function Close-OutlookPopup {
    param(
        [string]$titlePattern
    )

    $popup = Get-Process outlook -ErrorAction SilentlyContinue | Where-Object {
        $_.MainWindowTitle -match $titlePattern
    } | Select-Object -First 1

    if ($popup) {
        [WinAPI]::SetForegroundWindow($popup.MainWindowHandle) | Out-Null
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait("%{F4}")
        return $true
    }
    return $false
}

# Ví dụ dùng:
Close-OutlookPopup "New Profile"
Close-OutlookPopup "Welcome to Outlook"

function Format-SizeGB ($bytes) {
    return [math]::Round($bytes / 1GB, 2)
}

try {
    # Mở Outlook
    $outlookProcess = Start-Process "outlook.exe" -PassThru
    Start-Sleep -Seconds 3

    Close-OutlookFirstRun
    Start-Sleep -Seconds 5

    # Tạo COM object Outlook để lấy dữ liệu
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")

    $dungLuongChiTiet = @()
    $tongDungLuong = 0
    $tongToiDa = 50 * 1GB

    for ($i = 1; $i -le $namespace.Folders.Count; $i++) {
        $folder = $namespace.Folders.Item($i)
        $store = $folder.Store
        $filePath = $store.FilePath

        if ([string]::IsNullOrWhiteSpace($filePath) -or -not (Test-Path $filePath)) {
            continue
        }

        $file = Get-Item $filePath
        $fileSize = $file.Length
        $tongDungLuong += $fileSize

        $sizeGB = Format-SizeGB $fileSize
        $dungLuongChiTiet += "$($folder.Name): $sizeGB GB ($filePath)"
    }

    if ($dungLuongChiTiet.Count -gt 0) {
        $dungLuongDaDung = Format-SizeGB $tongDungLuong
        $dungLuongToiDa = Format-SizeGB $tongToiDa
        $outlookSizeStatus = "Tổng dung lượng Outlook: $dungLuongDaDung GB / $dungLuongToiDa GB`n" + ($dungLuongChiTiet -join "`n")
    }
    else {
        $outlookSizeStatus = "Không sử dụng Outlook"
    }

    # Tắt Outlook sau khi lấy dữ liệu
    Stop-Process -Name outlook -ErrorAction SilentlyContinue
}
catch {
    $outlookSizeStatus = "Không sử dụng Outlook"
}

# 13. Check file Server/Nas/Onedrive :
    function Check-FileSharingStatus {
        $sharedFolders = Get-SmbShare -Special $false -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "IPC$" }
        $networkDrives = Get-WmiObject -Class Win32_NetworkConnection -ErrorAction SilentlyContinue |
                        Select-Object -ExpandProperty RemoteName
        $output = @()

        if ($sharedFolders.Count -gt 0) {
            $output += "Chia sẻ thư mục: " + ($sharedFolders | Select-Object -ExpandProperty Name -Unique) -join ", "
        }

        if ($networkDrives.Count -gt 0) {
            $output += "Kết nối tới: " + ($networkDrives -join ", ")
        }

        if ($env:OneDrive) {
            $output += "Đang sử dụng OneDrive: $($env:OneDrive)"
        }

        if ($output.Count -eq 0) {
            return "Không có chia sẻ hoặc kết nối file từ xa"
        }

        return $output -join " | "
    }

    $fileShareStatus = Check-FileSharingStatus

# 14. Check Task Manager :
    $cpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
    $cpuUsage = [math]::Round($cpuUsage, 2)

    $ramUsage = (Get-Counter '\Memory\% Committed Bytes In Use').CounterSamples.CookedValue
    $ramUsage = [math]::Round($ramUsage, 2)

# 15. Check Windows Update :
    function Check-WindowsUpdateStatus {
        try {
            $wu = Get-Service wuauserv -ErrorAction Stop
            $uso = Get-Service UsoSvc -ErrorAction Stop

            # Kiểm tra nếu dịch vụ bị disable
            if ($wu.StartType -eq 'Disabled' -or $uso.StartType -eq 'Disabled') {
                return "Windows Update bị vô hiệu hóa (services bị Disabled)"
            }

            # Tiếp tục kiểm tra bản cập nhật
            $Session = New-Object -ComObject Microsoft.Update.Session
            $Searcher = $Session.CreateUpdateSearcher()
            $SearchResult = $Searcher.Search("IsInstalled=0 and Type='Software'")

            if ($SearchResult.Updates.Count -eq 0) {
                return "Hệ thống đã được cập nhật đầy đủ"
            } else {
                return "Có $($SearchResult.Updates.Count) bản cập nhật đang chờ cài đặt"
            }
        } catch {
            return "Không thể kiểm tra trạng thái cập nhật (có thể bị chặn bởi phần mềm ngoài)"
        }
    }     

    $updateStatus = Check-WindowsUpdateStatus

# 17. Clean up the computer :
    $tempPath = $env:TEMP
    $cleanupStatus = "Không xác định"

    try {
        Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        $cleanupStatus = "Đã làm sạch thư mục Temp"
    } catch {
        $cleanupStatus = "Xảy ra lỗi khi làm sạch thư mục Temp - $_"
    }

# 18. Scan virus :
    try {
        Start-MpScan -ScanType QuickScan
        $scanResult = "Đã Scan Virus bằng Windows Security"
    } catch {
        $scanResult = "Không thể Scan Virus"
    }

    Write-Output $scanResult

# 19. Virus soft expire date :
    function Get-AntivirusStatus {
        try {
            $avList = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction Stop
            $thirdPartyAVs = $avList | Where-Object { $_.displayName -notmatch "Defender" }

            if ($thirdPartyAVs) {
                return ($thirdPartyAVs | Select-Object -ExpandProperty displayName -Unique) -join ", "
            } else {
                return "Windows Defender"
            }
        } catch {
            return "Không thể kiểm tra phần mềm diệt virus (có thể bị chặn)"
        }
    }

    $antivirusStatus = Get-AntivirusStatus



    # Giao diện bảng kết quả Chức năng bảo trì
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $f = New-Object Windows.Forms.Form
    $f.Text = "Thông tin hệ thống"
    $f.Size = New-Object Drawing.Size(540,600)
    $f.StartPosition = "CenterScreen"

    $txt = New-Object Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.ScrollBars = "Vertical"
    $txt.Font = New-Object Drawing.Font("Segoe UI",11)
    $txt.Size = New-Object Drawing.Size(520,540)
    $txt.Location = New-Object Drawing.Point(10,10)
    $txt.Font = New-Object Drawing.Font("Consolas", 11)  # font đều khoảng cách

# Chuẩn hóa giá trị biến để luôn xuống dòng
$diskInfo     = "$diskInfo`r`n"
$freeStr      = "$freeStr`r`n"
$diskHealth   = "$diskHealth`r`n"
$batteryLife  = "$batteryLife`r`n"
$CheckNetwork = "$CheckNetwork`r`n"
$checkTime    = "$checkTime`r`n"
$officeStatus = "$officeStatus`r`n"
$outlookSizeStatus = "$outlookSizeStatus`r`n"
$fileShareStatus   = "$fileShareStatus`r`n"
$updateStatus = "$updateStatus`r`n"
$cleanupStatus = "$cleanupStatus`r`n"
$scanResult   = "$scanResult`r`n"
$antivirusStatus = "$antivirusStatus`r`n"

$txt.Text = @"
Check Tasks/User : $userName`r`n

1. Check CPU (Temperature) : Nhiệt độ trung bình $cpuTemp`r`n

2. Check Mainboard : $mainboard`r`n

3. Check RAM (DDR? & Bus?) : $ramSize, $ramType, $ramBus`r`n

4. HardDisk Total Space (SSD/HDD):
$diskInfo

5. HardDisk Free Space (SSD/HDD):
$freeStr

6. Check HardDisk Life:
$diskHealth

7. Check Battery Life Laptop : $batteryLife

8. Check Noisy : Bình thường`r`n

9. Check Network Connection : $CheckNetwork

10. Check Date & Time : $checkTime

11. Check Office status : $officeStatus

12. Check Mailbox status : $outlookSizeStatus

13. Check file Server/Nas/Onedrive : $fileShareStatus

14. Check Task Manager : CPU = $cpuUsage% | RAM = $ramUsage%`r`n

15. Check Windows Update : $updateStatus

16. Backup data : Chưa thực hiện`r`n

17. Clean up the computer : $cleanupStatus

18. Scan virus : $scanResult

19. Virus soft expire date : $antivirusStatus

20. Check Keyboard, Mouse : Bình thường
"@

    $f.Controls.Add($txt)

    $btn = New-Object Windows.Forms.Button
    $btn.Text    = "Đóng"
    $btn.Size    = New-Object Drawing.Size(100,30)
    $btn.Location = New-Object Drawing.Point(220,560)
    $btn.Add_Click({ $f.Close() })
    $f.Controls.Add($btn)

    $f.ShowDialog()
}





#------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------Hàm map link Github để tải Office--------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------
# 1. Biến $officeLinks
$officeLinks = @{
    "Office 2019" = "https://github.com/NNS-Dev02/MiT/releases/download/v1.0/Office.2019.zip"
    "Office 2021" = "https://github.com/NNS-Dev02/MiT/releases/download/v1.0/Office.2021.zip"
    "Office 2024" = "https://github.com/NNS-Dev02/MiT/releases/download/v1.0/Office.2024.zip"
    "Office 365"  = "https://github.com/NNS-Dev02/MiT/releases/download/v1.0/Office.365.zip"
}
# 2. Hàm chọn phiên bản Office
function Show-OfficeSuiteForm {
    param(
        [Parameter(Mandatory = $true)]
        [Hashtable]$Links
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Chọn phiên bản Office"
    $form.Size = New-Object System.Drawing.Size(350, 250)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.ControlBox = $false
    $form.Topmost = $true

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(20, 20)
    $checkedListBox.Size = New-Object System.Drawing.Size(300, 140)
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    foreach ($key in $Links.Keys | Sort-Object) {
        $checkedListBox.Items.Add($key)
    }
    $form.Controls.Add($checkedListBox)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(120, 170)
    $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $checkedItems = @()
        foreach ($item in $checkedListBox.CheckedItems) {
            $checkedItems += $item
        }
        return $checkedItems
    } else {
        return @()
    }
}
# 3. Xử lý tải và giải nén Office .rar
switch ($action) {
    "officeSuite" {
        $chs = Show-OfficeSuiteForm -Links $officeLinks 
        foreach ($v in $chs) {
            $url = $officeLinks[$v]
            $fileName = Split-Path $url -Leaf
            $outPath = Join-Path "C:\MiT" $fileName

            # Tải về file ZIP hoặc RAR
            Download-WithLoading -Url $url -OutFile $outPath -Message "Đang tải $v..."

            # Nếu là .zip thì giải nén
            if ($outPath -like "*.zip") {
                $extractPath = Join-Path "C:\MiT" ($v -replace '\s','')
                if (-not (Test-Path $extractPath)) {
                    New-Item -ItemType Directory -Path $extractPath | Out-Null
                }

                try {
                    Expand-Archive -Path $outPath -DestinationPath $extractPath -Force
                    Start-Process "$extractPath"
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Giải nén thất bại: $_","Lỗi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                Start-Process $outPath
            }
        }
    }

    default {
        Write-Host "Hành động không xác định: $action"
    }
}
# 4. Hàm Download-WithLoading
function Download-WithLoading {
    param (
        [string]$Url,
        [string]$OutFile,
        [string]$Message = "Đang tải..."
    )

    $form = New-Object Windows.Forms.Form
    $form.Text = "Đang tải"
    $form.Size = New-Object Drawing.Size(400, 100)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.ControlBox = $false
    $form.Topmost = $true

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleCenter"
    $label.Font = New-Object Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($label)

    $job = Start-Job {
        param($u, $o)
        Invoke-WebRequest -Uri $u -OutFile $o -UseBasicParsing
    } -ArgumentList $Url, $OutFile

    while ($job.State -eq "Running") {
        $form.Refresh()
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
    }

    $form.Close()
    Receive-Job $job | Out-Null
    Remove-Job $job
}






#------------------------------------------------------------------------------------------------------------------------------------
#-------------------------------------- Hàm map link Google Drive để tải App Crack --------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------

# 1. Biến $appCrack
$appCrack = @{
    "Adobe Acrobat Pro" = "https://drive.google.com/uc?export=download&id=1cYVs3Jnw5S9WsBNau3BRie7XaEDPZMhN"
    "AutoCAD 2022"       = "https://drive.google.com/uc?export=download&id=1xODjUqtQN4f9SY4U30Ff-ghsrSIeDb34"
}

# 2. Hàm chọn ứng dụng App Crack
function Show-appCrackSuiteForm {
    param(
        [Parameter(Mandatory = $true)]
        [Hashtable]$Links
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Chọn ứng dụng App Crack"
    $form.Size = New-Object System.Drawing.Size(350, 250)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.ControlBox = $false
    $form.Topmost = $true

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(20, 20)
    $checkedListBox.Size = New-Object System.Drawing.Size(300, 140)
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    foreach ($key in $Links.Keys | Sort-Object) {
        $checkedListBox.Items.Add($key)
    }
    $form.Controls.Add($checkedListBox)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(120, 170)
    $btnOk.Size = New-Object System.Drawing.Size(100, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $checkedItems = @()
        foreach ($item in $checkedListBox.CheckedItems) {
            $checkedItems += $item
        }
        return $checkedItems
    } else {
        return @()
    }
}

# 3. Xử lý tải và giải nén App Crack
switch ($action) {
    "appCrackSuite" {
        $chs = Show-appCrackSuiteForm -Links $appCrack
        foreach ($v in $chs) {
            $url = $appCrack[$v]
            $fileName = "$v.rar"  # hoặc .zip nếu bạn dùng định dạng zip
            $outPath = Join-Path "C:\MiT" $fileName

            # Tải về file
            Download-WithLoading -Url $url -OutFile $outPath -Message "Đang tải $v..."

            # Nếu là .zip thì giải nén
            if ($outPath -like "*.zip") {
                $extractPath = Join-Path "C:\MiT" ($v -replace '\s','')
                if (-not (Test-Path $extractPath)) {
                    New-Item -ItemType Directory -Path $extractPath | Out-Null
                }

                try {
                    Expand-Archive -Path $outPath -DestinationPath $extractPath -Force
                    Start-Process "$extractPath"
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Giải nén thất bại: $_","Lỗi",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
                }
            } else {
                Start-Process $outPath
            }
        }
    }

    default {
        Write-Host "Hành động không xác định: $action"
    }
}

# 4. Hàm Download-WithLoading
function Download-WithLoading {
    param (
        [string]$Url,
        [string]$OutFile,
        [string]$Message = "Đang tải..."
    )

    $form = New-Object Windows.Forms.Form
    $form.Text = "Đang tải"
    $form.Size = New-Object Drawing.Size(400, 100)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.ControlBox = $false
    $form.Topmost = $true

    $label = New-Object Windows.Forms.Label
    $label.Text = $Message
    $label.Dock = "Fill"
    $label.TextAlign = "MiddleCenter"
    $label.Font = New-Object Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($label)

    $job = Start-Job {
        param($u, $o)
        Invoke-WebRequest -Uri $u -OutFile $o -UseBasicParsing
    } -ArgumentList $Url, $OutFile

    while ($job.State -eq "Running") {
        $form.Refresh()
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()
    }

    $form.Close()
    Receive-Job $job | Out-Null
    Remove-Job $job
}






#------------------------------------------------------------------------------------------------------------------------------------
#--------------------------------------------------Hàm map link Github tải Ứng dụng--------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------
# Hàm cài ứng dụng đã chọn
function Install-SelectedApplications {
    $sel = $listBox.CheckedItems
    if ($sel.Count -eq 0) {
        [Windows.Forms.MessageBox]::Show("Vui lòng chọn ít nhất một ứng dụng.","Thông báo",0,0)
        return
    }

    # Thiết lập đường dẫn và tải file apps.txt nếu cần
    $folder   = "C:\MiT"
    $appsUrl  = "https://raw.githubusercontent.com/NNS-Dev02/MiT/main/apps.txt"
    $appsPath = Join-Path $folder "apps.txt"

    if (-not (Test-Path $folder)) { New-Item $folder -ItemType Directory | Out-Null }
    if (-not (Test-Path $appsPath)) {
        Download-WithLoading -Url $appsUrl -OutFile $appsPath -Message "Đang tải apps.txt..."
    }

    # Đọc apps.txt
    $script:appLines = Get-Content $appsPath | Where-Object { $_ -match "\|" }
    $script:menuItems = foreach ($l in $script:appLines) {
        if ($l -match "^\s*(.+?)\s*\|\s*(.+?)\s*$") {
            [PSCustomObject]@{ Name = $matches[1].Trim(); Value = $matches[2].Trim() }
        }
    }

    foreach ($name in $sel) {
        $entry = $script:menuItems | Where-Object { $_.Name -eq $name }
        $val   = $entry.Value
        $cleanName = $name -replace '^\d+\.\s*',''

        switch ($val) {
            "script" {
                irm https://get.activated.win | iex
            }
            "officeSuite" {
                $chs = Show-OfficeSuiteForm -Links $officeLinks 
                foreach ($v in $chs) { Start-Process $officeLinks[$v] }
            }
            "appCrackSuite" {
                $chs = Show-appCrackSuiteForm -Links $appCrack
                foreach ($v in $chs) {
                    Start-Process $appCrack[$v]
                }
            }
            "maintenance" {
                $confirmationForm = New-Object System.Windows.Forms.Form
                $confirmationForm.Text = "Thông báo"
                $confirmationForm.Size = New-Object System.Drawing.Size(450, 180)
                $confirmationForm.StartPosition = "CenterScreen"
                $confirmationForm.FormBorderStyle = "FixedDialog"
                $confirmationForm.MinimizeBox = $false
                $confirmationForm.MaximizeBox = $false
                $confirmationForm.ControlBox = $false 

                $label = New-Object System.Windows.Forms.Label
                $label.Text = "Bạn đã cài ứng dụng CPU-Z, Core Temp, CrystalDiskInfo chưa?"
                $label.Location = New-Object System.Drawing.Point(20, 30) 
                $label.Size = New-Object System.Drawing.Size(400, 60) 
                $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter 
                $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                $confirmationForm.Controls.Add($label)

                $btnRoi = New-Object System.Windows.Forms.Button
                $btnRoi.Text = "Rồi"
                $btnRoi.Location = New-Object System.Drawing.Point(100, 100) 
                $btnRoi.Size = New-Object System.Drawing.Size(75, 30)
                $btnRoi.DialogResult = [System.Windows.Forms.DialogResult]::Yes 
                $confirmationForm.Controls.Add($btnRoi)

                $btnChua = New-Object System.Windows.Forms.Button
                $btnChua.Text = "Chưa"
                $btnChua.Location = New-Object System.Drawing.Point(250, 100) 
                $btnChua.Size = New-Object System.Drawing.Size(75, 30)
                $btnChua.DialogResult = [System.Windows.Forms.DialogResult]::No 
                $confirmationForm.Controls.Add($btnChua)

                $dialogResult = $confirmationForm.ShowDialog()

                if ($dialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Show-SystemInfoForm
                } else {
                    $alertForm = New-Object System.Windows.Forms.Form
                    $alertForm.Text = "Thông báo"
                    $alertForm.Size = New-Object System.Drawing.Size(450, 180) 
                    $alertForm.StartPosition = "CenterScreen"
                    $alertForm.FormBorderStyle = "FixedDialog"
                    $alertForm.MinimizeBox = $false
                    $alertForm.MaximizeBox = $false
                    $alertForm.ControlBox = $false 

                    $alertLabel = New-Object System.Windows.Forms.Label
                    $alertLabel.Text = "Hãy cài ứng ứng dụng CPU-Z, Core Temp, CrystalDiskInfo ở mục 10,11,12!"
                    $alertLabel.Location = New-Object System.Drawing.Point(20, 30)
                    $alertLabel.Size = New-Object System.Drawing.Size(400, 60) 
                    $alertLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter 
                    $alertLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
                    $alertForm.Controls.Add($alertLabel)

                    $btnDongY = New-Object System.Windows.Forms.Button
                    $btnDongY.Text = "Đồng ý"
                    $btnDongY.Location = New-Object System.Drawing.Point(([int](($alertForm.ClientSize.Width - 75) / 2)), 100)
                    $btnDongY.Size = New-Object System.Drawing.Size(75, 30)
                    $btnDongY.DialogResult = [System.Windows.Forms.DialogResult]::OK 
                    $alertForm.Controls.Add($btnDongY)

                    $alertForm.ShowDialog()
                }
            }
            default {
            # Xử lý app thông thường qua URL
            $downloadUrl = $val
            $fileName = Split-Path $downloadUrl -Leaf
            $outPath = Join-Path $folder $fileName

            # Kiểm tra nếu file đã tồn tại
            if (Test-Path $outPath) {
                $reDownload = [Windows.Forms.MessageBox]::Show(
                    "Tệp '$fileName' đã tồn tại. Bạn có muốn tải lại không?", 
                    "Tệp đã tồn tại", 
                    [Windows.Forms.MessageBoxButtons]::YesNo, 
                    [Windows.Forms.MessageBoxIcon]::Question
                )

                if ($reDownload -eq [Windows.Forms.DialogResult]::Yes) {
                    Download-WithLoading -Url $downloadUrl -OutFile $outPath -Message "Đang tải $fileName..."
                }
                elseif ($reDownload -eq [Windows.Forms.DialogResult]::No) {
                    # Bỏ qua ứng dụng này
                    continue
                }
            }
            else {
                Download-WithLoading -Url $downloadUrl -OutFile $outPath -Message "Đang tải $fileName..."
            }

            # Chỉ cài nếu file tồn tại (nghĩa là không bấm No)
            if (Test-Path $outPath) {
                Start-Process $outPath
            }
        }
        }
    }
}





#--------------------------------------------------------------------------------------------------------------
#--------------------------------------------------Form Chính--------------------------------------------------
#--------------------------------------------------------------------------------------------------------------
# 1. Tạo form chính
    $form = New-Object Windows.Forms.Form
    $form.Text           = "Công Ty TNHH Thương Mại Và Sản Xuất MiT"
    $form.Size           = New-Object Drawing.Size(600,650)
    $form.StartPosition = "CenterScreen"
    $form.Font           = New-Object Drawing.Font("Segoe UI",12)

# 2. Logo và Apps.txt
    $pic = New-Object Windows.Forms.PictureBox
    $pic.Size     = New-Object Drawing.Size(200,150)
    $pic.SizeMode = "StretchImage"
    $pic.Location = New-Object Drawing.Point(([int](($form.ClientSize.Width-200)/2)),10)

    $folder   = "C:\MiT"
    $imgUrl   = "https://raw.githubusercontent.com/NNS-Dev02/MiT/main/logo.png"
    $appsUrl  = "https://raw.githubusercontent.com/NNS-Dev02/MiT/main/apps.txt"
    $imgPath  = Join-Path $folder "logo.png"

    if (-not (Test-Path $folder)) { New-Item $folder -ItemType Directory | Out-Null }
    if (-not (Test-Path $imgPath)) { Download-WithLoading -Url $imgUrl -OutFile $imgPath -Message "Đang tải logo..." }

    if (Test-Path $imgPath) {
        try { $pic.Image = [Drawing.Image]::FromFile($imgPath) } catch {}
    }
    $form.Controls.Add($pic)

# 3. CheckedListBox ứng dụng
    $listBox = New-Object Windows.Forms.CheckedListBox
    $listBox.Size          = New-Object Drawing.Size(560,350)
    $listBox.Location      = New-Object Drawing.Point(10,180)
    $listBox.CheckOnClick  = $true
    $listBox.Font          = New-Object Drawing.Font("Segoe UI",13)

    $appsPath = Join-Path $folder "apps.txt"
    if (-not (Test-Path $appsPath)) { Download-WithLoading -Url $appsUrl -OutFile $appsPath -Message "Đang tải apps.txt..." }
    $appLines = Get-Content $appsPath | Where-Object { $_ -match "\|" }
    $menuItems = foreach ($l in $appLines) {
        if ($l -match "^\s*(.+?)\s*\|\s*(.+?)\s*$") {
            [PSCustomObject]@{ Name = $matches[1].Trim(); Value = $matches[2].Trim() }
        }
    }
    $menuItems.ForEach({ $listBox.Items.Add($_.Name) })
    $form.Controls.Add($listBox)

# 4. Nút Cài đặt
    $btnInstall = New-Object Windows.Forms.Button
    $btnInstall.Text     = "Cài đặt"
    $btnInstall.Size     = New-Object Drawing.Size(80,30)
    $btnInstall.Location = New-Object Drawing.Point(350,550)
    $btnInstall.Font     = New-Object Drawing.Font("Segoe UI",11)
    $btnInstall.Add_Click({ Install-SelectedApplications })
    $form.Controls.Add($btnInstall)

# 5. Nút Thoát
    $btnExit = New-Object Windows.Forms.Button
    $btnExit.Text     = "Thoát"
    $btnExit.Size     = New-Object Drawing.Size(80,30)
    $btnExit.Location = New-Object Drawing.Point(460,550)
    $btnExit.Font     = New-Object Drawing.Font("Segoe UI",11)
    $btnExit.Add_Click({ $form.Close() })
    $form.Controls.Add($btnExit)


$form.ShowDialog()