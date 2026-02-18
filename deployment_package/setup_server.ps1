# SERVER SETUP CONFIGURATION
# Quick setup script for server deployment

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PLZ CV ENGINE - SERVER SETUP" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Get network path
Write-Host "Step 1: Configure Network Path" -ForegroundColor Yellow
Write-Host ""
Write-Host "Enter the network path where the application should monitor:"
Write-Host "Examples:"
Write-Host "  \\SERVER\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers"
Write-Host "  D:\SharedData\Premium Leaf Zimbabwe\v2SavannahTEST - Collection Vouchers"
Write-Host ""
$networkPath = Read-Host "Network Path"

if ([string]::IsNullOrWhiteSpace($networkPath)) {
    Write-Host "Error: Path cannot be empty" -ForegroundColor Red
    exit 1
}

# Step 2: Test path accessibility
Write-Host ""
Write-Host "Step 2: Testing Path Accessibility..." -ForegroundColor Yellow
if (Test-Path $networkPath) {
    Write-Host "✓ Path is accessible" -ForegroundColor Green
} else {
    Write-Host "✗ Path is NOT accessible" -ForegroundColor Red
    $create = Read-Host "Do you want to create it? (yes/no)"
    if ($create -eq "yes") {
        try {
            New-Item -Path $networkPath -ItemType Directory -Force | Out-Null
            Write-Host "✓ Path created" -ForegroundColor Green
        } catch {
            Write-Host "✗ Failed to create path: $($_.Exception.Message)" -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host "Please create the path manually and run this script again." -ForegroundColor Yellow
        exit 1
    }
}

# Step 3: Create folder structure
Write-Host ""
Write-Host "Step 3: Creating Folder Structure..." -ForegroundColor Yellow
$folders = @(
    "$networkPath\Json",
    "$networkPath\Templates",
    "$networkPath\Populated Template",
    "$networkPath\PDF CVs"
)

foreach ($folder in $folders) {
    if (!(Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
        Write-Host "  ✓ Created: $folder" -ForegroundColor Green
    } else {
        Write-Host "  ✓ Exists: $folder" -ForegroundColor Gray
    }
}

# Step 4: Update config.ini
Write-Host ""
Write-Host "Step 4: Updating config.ini..." -ForegroundColor Yellow

$configPath = ".\deployment_package\config.ini"
if (!(Test-Path $configPath)) {
    $configPath = ".\config.ini"
}

if (Test-Path $configPath) {
    $config = Get-Content $configPath -Raw
    
    # Update root_path
    if ($config -match "root_path\s*=\s*.*") {
        $config = $config -replace "root_path\s*=\s*.*", "root_path = $networkPath"
    } else {
        # Add root_path if not exists
        $config = $config -replace "\[PATHS\]", "[PATHS]`nroot_path = $networkPath"
    }
    
    Set-Content $configPath -Value $config
    Write-Host "✓ config.ini updated" -ForegroundColor Green
} else {
    Write-Host "✗ config.ini not found" -ForegroundColor Red
}

# Step 5: Check for Excel
Write-Host ""
Write-Host "Step 5: Checking for Excel Installation..." -ForegroundColor Yellow
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "✓ Microsoft Excel is installed" -ForegroundColor Green
} catch {
    Write-Host "✗ Microsoft Excel not found" -ForegroundColor Red
    Write-Host "  Please install Microsoft Excel on this server" -ForegroundColor Yellow
}

# Step 6: Check for template
Write-Host ""
Write-Host "Step 6: Checking for Template File..." -ForegroundColor Yellow
$templatePath = "$networkPath\Templates\Collection Voucher Template.xlsx"
if (Test-Path $templatePath) {
    Write-Host "✓ Template file exists" -ForegroundColor Green
} else {
    Write-Host "✗ Template file not found" -ForegroundColor Red
    Write-Host "  Please copy 'Collection Voucher Template.xlsx' to:" -ForegroundColor Yellow
    Write-Host "  $networkPath\Templates\" -ForegroundColor Yellow
}

# Step 7: Email configuration
Write-Host ""
Write-Host "Step 7: Email Configuration" -ForegroundColor Yellow
Write-Host ""
Write-Host "Do you want to configure email settings now? (yes/no)"
$configEmail = Read-Host

if ($configEmail -eq "yes") {
    Write-Host ""
    Write-Host "Email Configuration:" -ForegroundColor Cyan
    
    $enableEmail = Read-Host "Enable email notifications? (true/false)"
    if ($enableEmail -eq "true") {
        $smtpServer = Read-Host "SMTP Server (e.g., smtp.gmail.com)"
        $smtpPort = Read-Host "SMTP Port (default: 587)"
        $senderEmail = Read-Host "Sender Email"
        $senderPassword = Read-Host "Sender Password/App Password" -AsSecureString
        $senderPasswordPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($senderPassword))
        $recipients = Read-Host "Recipients (comma-separated)"
        
        # Update config with email settings
        if (Test-Path $configPath) {
            $config = Get-Content $configPath -Raw
            $config = $config -replace "enabled\s*=\s*.*", "enabled = $enableEmail"
            $config = $config -replace "smtp_server\s*=\s*.*", "smtp_server = $smtpServer"
            $config = $config -replace "smtp_port\s*=\s*.*", "smtp_port = $smtpPort"
            $config = $config -replace "sender_email\s*=\s*.*", "sender_email = $senderEmail"
            $config = $config -replace "sender_password\s*=\s*.*", "sender_password = $senderPasswordPlain"
            $config = $config -replace "recipients\s*=\s*.*", "recipients = $recipients"
            
            Set-Content $configPath -Value $config
            Write-Host "✓ Email configuration saved" -ForegroundColor Green
        }
    }
}

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SETUP COMPLETE!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Configuration Summary:" -ForegroundColor Yellow
Write-Host "  Monitoring Path: $networkPath"
Write-Host "  Config File: $configPath"
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Copy 'Collection Voucher Template.xlsx' to Templates folder"
Write-Host "  2. Test the application: .\PLZ_CV_Engine.exe"
Write-Host "  3. Set up as Windows Service (see SERVER_DEPLOYMENT.md)"
Write-Host "  4. Update Power Automate to use new path"
Write-Host ""
Write-Host "For detailed instructions, see SERVER_DEPLOYMENT.md" -ForegroundColor Cyan
Write-Host ""
