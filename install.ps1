$ErrorActionPreference = 'Stop'

# UNIFIED SOLUTIONS CONFIGURATION
$rustdesk_server = "unified-rustdesk.synology.me"
$rustdesk_key = "MpyjeId83qkVEy+QwrGJNdBDbUCN0uNjwCyWtHaKQ8E="

Write-Host "🔧 Unified Solutions - Emergency Remote Access Setup" -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan

# Check administrator privileges
if (-Not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "❌ ERROR: Must run as Administrator!" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    Exit 1
}

try {
    # Create temporary directory
    $tempPath = "$env:TEMP\UnifiedRustDesk"
    if (Test-Path $tempPath) { Remove-Item $tempPath -Recurse -Force }
    New-Item -ItemType Directory -Path $tempPath | Out-Null
    Set-Location $tempPath

    Write-Host "📡 Downloading latest RustDesk..." -ForegroundColor Green
    
    # Get latest release from GitHub API
    $apiResponse = Invoke-RestMethod "https://api.github.com/repos/rustdesk/rustdesk/releases/latest"
    $asset = $apiResponse.assets | Where-Object { $_.name -match "rustdesk.*x86_64.*\.exe$" -and $_.name -notmatch "sciter" } | Select-Object -First 1
    
    if (!$asset) {
        throw "Could not find Windows installer"
    }
    
    $downloadUrl = $asset.browser_download_url
    $fileName = $asset.name
    $version = $apiResponse.tag_name
    
    Write-Host "   Version: $version" -ForegroundColor Gray
    Write-Host "   File: $fileName" -ForegroundColor Gray
    
    # Download installer
    Invoke-WebRequest -Uri $downloadUrl -OutFile $fileName -UseBasicParsing
    
    Write-Host "🚀 Installing RustDesk..." -ForegroundColor Green
    Start-Process ".\$fileName" -ArgumentList "--silent-install" -Wait
    
    Start-Sleep -Seconds 3
    
    # Find installation path
    $rustdeskExe = @(
        "${env:ProgramFiles}\RustDesk\rustdesk.exe",
        "${env:ProgramFiles(x86)}\RustDesk\rustdesk.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (!$rustdeskExe) {
        throw "RustDesk installation not found"
    }
    
    Write-Host "⚙️ Configuring server connection..." -ForegroundColor Green
    
    # Build configuration
    $config = @{
        host = $rustdesk_server
        key = $rustdesk_key
    } | ConvertTo-Json -Compress
    $configBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($config))
    
    # Apply configuration
    & $rustdeskExe --config $configBase64
    
    # Generate secure password
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    $password = -join (1..12 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    
    # Set password
    & $rustdeskExe --password $password
    
    # Install as service
    & $rustdeskExe --install-service
    
    Start-Sleep -Seconds 2
    
    # Get device ID
    $deviceId = & $rustdeskExe --get-id
    
    # Display results
    Clear-Host
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "✅ INSTALLATION SUCCESSFUL!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "🏢 Business: $env:COMPUTERNAME" -ForegroundColor Yellow
    Write-Host "🆔 Device ID: $deviceId" -ForegroundColor Yellow
    Write-Host "🔐 Password: $password" -ForegroundColor Yellow
    Write-Host "🌐 Server: $rustdesk_server" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "📝 TECHNICIAN: Save this information!" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Green
    
    # Save to desktop
    $infoFile = "$env:USERPROFILE\Desktop\Unified-Remote-Access.txt"
    @"
UNIFIED SOLUTIONS - Remote Access Information
Installed: $(Get-Date -Format "yyyy-MM-dd HH:mm")

Business Computer: $env:COMPUTERNAME
Device ID: $deviceId
Access Password: $password
Server: $rustdesk_server

--- FOR TECHNICIAN RECORDS ---
This system now has emergency remote access configured.
"@ | Out-File $infoFile -Encoding UTF8
    
    Write-Host "💾 Information saved to desktop: Unified-Remote-Access.txt" -ForegroundColor Green
    
} catch {
    Write-Host "❌ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    # Cleanup
    if (Test-Path $tempPath) {
        Set-Location $env:TEMP
        Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "Press any key to close..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
