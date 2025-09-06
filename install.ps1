$ErrorActionPreference = 'Stop'

# UNIFIED SOLUTIONS CONFIGURATION
$rustdesk_server = "unified-rustdesk.synology.me"
$rustdesk_key = "MpyjeId83qkVEy+QwrGJNdBDbUCN0uNjwCyWtHaKQ8E="

Write-Host "🔧 Unified Solutions - Emergency Remote Access Setup" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "Target System: $env:COMPUTERNAME" -ForegroundColor Gray
Write-Host "Technician: $env:USERNAME" -ForegroundColor Gray
Write-Host "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "====================================================" -ForegroundColor Cyan

# Verify administrator privileges
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ ERROR: Administrator privileges required!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    Exit 1
}

# Technician confirmation
Write-Host "⚠️  TECHNICIAN CONFIRMATION REQUIRED" -ForegroundColor Yellow
Write-Host "This will install RustDesk remote access software on this client system." -ForegroundColor White
$confirmation = Read-Host "Continue with installation? (Type YES to confirm)"

if ($confirmation -ne "YES") {
    Write-Host "❌ Installation cancelled by technician" -ForegroundColor Red
    Exit 1
}

try {
    # Create temporary workspace
    $tempPath = "$env:TEMP\UnifiedRustDesk-$(Get-Random)"
    New-Item -ItemType Directory -Path $tempPath | Out-Null
    Set-Location $tempPath

    Write-Host "📡 Downloading latest RustDesk version..." -ForegroundColor Green
    
    # Get latest release from GitHub API
    $apiResponse = Invoke-RestMethod "https://api.github.com/repos/rustdesk/rustdesk/releases/latest" -UseBasicParsing
    $asset = $apiResponse.assets | Where-Object { 
        $_.name -match "rustdesk.*x86_64.*\.exe$" -and 
        $_.name -notmatch "sciter" 
    } | Select-Object -First 1
    
    if (!$asset) {
        throw "Could not find Windows x64 installer"
    }
    
    $downloadUrl = $asset.browser_download_url
    $fileName = $asset.name
    $version = $apiResponse.tag_name
    
    Write-Host "   Version: $version" -ForegroundColor Gray
    Write-Host "   File: $fileName" -ForegroundColor Gray
    
    # Download installer
    Invoke-WebRequest -Uri $downloadUrl -OutFile $fileName -UseBasicParsing
    
    Write-Host "🚀 Installing RustDesk..." -ForegroundColor Green
    Start-Process ".\$fileName" -ArgumentList "--silent-install" -Wait -NoNewWindow
    
    Start-Sleep -Seconds 5
    
    # Find RustDesk installation
    $possiblePaths = @(
        "${env:ProgramFiles}\RustDesk\rustdesk.exe",
        "${env:ProgramFiles(x86)}\RustDesk\rustdesk.exe",
        "${env:LOCALAPPDATA}\Programs\RustDesk\rustdesk.exe"
    )
    
    $rustdeskExe = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (!$rustdeskExe) {
        throw "RustDesk installation not found"
    }
    
    Write-Host "⚙️ Configuring Unified Solutions server..." -ForegroundColor Green
    
    # Build configuration
    $config = @{
        host = $rustdesk_server
        key = $rustdesk_key
    } | ConvertTo-Json -Compress
    
    $configBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($config))
    
    # Apply configuration
    & $rustdeskExe --config $configBase64
    
    Write-Host "🔐 Generating secure access password..." -ForegroundColor Green
    
    # Generate strong password
    $charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    $password = -join (1..14 | ForEach-Object { $charset[(Get-Random -Maximum $charset.Length)] })
    
    # Set password
    & $rustdeskExe --password $password
    
    Write-Host "🛠️ Installing system service..." -ForegroundColor Green
    & $rustdeskExe --install-service
    
    Start-Sleep -Seconds 3
    
    # Get device ID
    $deviceId = & $rustdeskExe --get-id
    
    if (!$deviceId) {
        throw "Could not retrieve device ID"
    }
    
    # Display results
    Clear-Host
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✅ INSTALLATION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "🏢 Client System: " -NoNewline -ForegroundColor White
    Write-Host "$env:COMPUTERNAME" -ForegroundColor Yellow
    Write-Host "🆔 Device ID: " -NoNewline -ForegroundColor White
    Write-Host "$deviceId" -ForegroundColor Yellow
    Write-Host "🔐 Access Password: " -NoNewline -ForegroundColor White
    Write-Host "$password" -ForegroundColor Yellow
    Write-Host "🌐 Server: " -NoNewline -ForegroundColor White
    Write-Host "$rustdesk_server" -ForegroundColor Yellow
    Write-Host "📅 Installed: " -NoNewline -ForegroundColor White
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "📋 TECHNICIAN: Record this information" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Green
    
    # Save client info to desktop
    $infoFile = "$env:PUBLIC\Desktop\Unified-Remote-Support-Info.txt"
    $clientInfo = @"
UNIFIED SOLUTIONS - Remote Support Access
Installation Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Your system now has emergency remote support capability installed.

System: $env:COMPUTERNAME
Support ID: $deviceId
Installation: $(Get-Date -Format "yyyy-MM-dd")

Contact Information:
Unified Solutions
Email: info@unified-solutions.ca
Location: Montreal, QC, Canada
"@
    
    $clientInfo | Out-File $infoFile -Encoding UTF8
    
    # Save technician record
    $techFile = "$env:USERPROFILE\Desktop\TECH-RECORD-$env:COMPUTERNAME.txt"
    $techRecord = @"
UNIFIED SOLUTIONS - TECHNICIAN RECORD
====================================

Client: $env:COMPUTERNAME
Device ID: $deviceId
Password: $password
Server: $rustdesk_server
Installed: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Technician: $env:USERNAME
Version: $version

NEXT STEPS:
1. Test remote connection
2. Add to client database
3. Inform client about support capability
====================================
"@
    
    $techRecord | Out-File $techFile -Encoding UTF8
    
    Write-Host "💾 Records saved to desktop" -ForegroundColor Green
    Write-Host "🧪 Test connection before leaving client site!" -ForegroundColor Yellow
    
} catch {
    Write-Host "❌ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Contact Unified Solutions for assistance" -ForegroundColor Yellow
} finally {
    # Cleanup
    if (Test-Path $tempPath) {
        Set-Location $env:TEMP
        Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Read-Host "Press Enter to close"
