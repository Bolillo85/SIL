$ErrorActionPreference = 'Stop'

$Base = 'C:\ProgramData\CuriosityCloud\Inventory'
$Output = Join-Path $Base 'Results'

New-Item -Path $Output -ItemType Directory -Force |
    Out-Null

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$Computer = $env:COMPUTERNAME

$Results = New-Object System.Collections.Generic.List[object]

function Add-InventoryItem {
    param(
        [string]$Category,
        [string]$Product,
        [string]$Edition = '',
        [string]$Version = '',
        [string]$Evidence = '',
        [string]$LicenseMetric = 'Review'
    )

    $Results.Add([PSCustomObject]@{
        ComputerName  = $Computer
        CollectedAt   = (Get-Date).ToString('s')
        Category      = $Category
        Product       = $Product
        Edition       = $Edition
        Version       = $Version
        Evidence      = $Evidence
        LicenseMetric = $LicenseMetric
    })
}

# Windows Server / sistema operativo
$OS = Get-CimInstance Win32_OperatingSystem
$CPU = Get-CimInstance Win32_Processor |
    Measure-Object -Property NumberOfLogicalProcessors -Sum

$VCPU = $CPU.Sum

Add-InventoryItem `
    -Category 'OperatingSystem' `
    -Product $OS.Caption `
    -Version $OS.Version `
    -Evidence "vCPU=$VCPU; Build=$($OS.BuildNumber)" `
    -LicenseMetric 'Core-Review'

# Software Microsoft instalado
$RegistryPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

$Software = @(
    Get-ItemProperty -Path $RegistryPaths -ErrorAction SilentlyContinue |
    Where-Object {
        $_.DisplayName -and (
            $_.DisplayName -match 'SQL Server|Exchange Server|SharePoint|Microsoft Office|Microsoft 365 Apps|System Center|BizTalk|Dynamics|Visual Studio|Project|Visio'
        )
    } |
    Sort-Object DisplayName, DisplayVersion -Unique
)

foreach ($App in $Software) {
    Add-InventoryItem `
        -Category 'InstalledSoftware' `
        -Product $App.DisplayName `
        -Version $App.DisplayVersion `
        -Evidence "Registry=$($App.PSPath)" `
        -LicenseMetric 'Review'
}

# Instancias de SQL Server
$SQLPath = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'

if (Test-Path $SQLPath) {
    $Instances = Get-ItemProperty $SQLPath

    foreach ($Property in $Instances.PSObject.Properties) {
        if ($Property.Name -match '^PS') { continue }

        $InstanceName = $Property.Name
        $InstanceID = $Property.Value

        $SetupPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$InstanceID\Setup"
        $Setup = Get-ItemProperty $SetupPath -ErrorAction SilentlyContinue

        Add-InventoryItem `
            -Category 'SQLInstance' `
            -Product "SQL Server - $InstanceName" `
            -Edition $Setup.Edition `
            -Version $Setup.Version `
            -Evidence "InstanceID=$InstanceID" `
            -LicenseMetric 'Core-or-SAL-Review'
    }
}

# Roles de Windows Server
$GetFeature = Get-Command Get-WindowsFeature -ErrorAction SilentlyContinue

if ($GetFeature) {
    $Features = Get-WindowsFeature |
        Where-Object {
            $_.Installed -and (
                $_.Name -match 'RDS|AD-Domain-Services|DNS|DHCP|Web-Server|File-Services|Hyper-V'
            )
        }

    foreach ($Feature in $Features) {
        $Metric = 'Role-Review'

        if ($Feature.Name -match '^RDS') {
            $Metric = 'Potential-SAL'
        }

        Add-InventoryItem `
            -Category 'WindowsRole' `
            -Product $Feature.DisplayName `
            -Evidence "FeatureName=$($Feature.Name)" `
            -LicenseMetric $Metric
    }
}

# Licencias RDS instaladas
try {
    $RDSFeature = Get-WindowsFeature -Name RDS-Licensing -ErrorAction SilentlyContinue

    if ($RDSFeature -and $RDSFeature.Installed) {

        $RDSPacks = @(
            Get-CimInstance `
                -Namespace 'root/CIMV2' `
                -ClassName 'Win32_TSLicenseKeyPack' `
                -ErrorAction Stop |
            Where-Object {
                $_.ProductVersion -match 'Windows Server' -and
                $_.TypeAndModel -match 'RDS'
            }
        )

        foreach ($Pack in $RDSPacks) {
            Add-InventoryItem `
                -Category 'RDSLicensePack' `
                -Product $Pack.ProductVersion `
                -Edition $Pack.TypeAndModel `
                -Evidence "KeyPackId=$($Pack.KeyPackId); TotalLicenses=$($Pack.TotalLicenses); AvailableLicenses=$($Pack.AvailableLicenses)" `
                -LicenseMetric 'RDS-CAL-Inventory'
        }
    }
}
catch {
    Write-Warning "RDS license inventory unavailable: $($_.Exception.Message)"
}
# Microsoft Office Click-to-Run
$OfficePath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'

$Office = Get-ItemProperty -Path $OfficePath -ErrorAction SilentlyContinue

if ($Office -and $Office.ProductReleaseIds) {
    foreach ($ProductID in ($Office.ProductReleaseIds -split ',')) {
        $ProductID = $ProductID.Trim()

        if ($ProductID -eq 'O365BusinessRetail') {
            $ProductName = 'Microsoft 365 Apps for business'
        }
        else {
            $ProductName = "Microsoft Office Click-to-Run ($ProductID)"
        }

        Add-InventoryItem `
            -Category 'OfficeClickToRun' `
            -Product $ProductName `
            -Edition $ProductID `
            -Version $Office.VersionToReport `
            -Evidence "Platform=$($Office.Platform); Path=$($Office.InstallationPath)" `
            -LicenseMetric 'License-Review'
    }
}
# Exportación
$CsvPath = Join-Path $Output "$Computer`_SPLA_$Timestamp.csv"

$Results |
    Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Inventory completed: $CsvPath"
Write-Host "Records: $($Results.Count)"


# Envio centralizado CuriosityCloud
try {
    $Archivo = Get-ChildItem 'C:\ProgramData\CuriosityCloud\Inventory\Results' -Filter '*_SPLA_*.csv' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $Archivo) {
        throw 'Inventory CSV not found'
    }

    $Respuesta = Invoke-WebRequest `
        -Uri 'http://sil.curiositycloud.com:8085/inventory' `
        -Method POST `
        -InFile $Archivo.FullName `
        -ContentType 'text/csv; charset=utf-8' `
        -UseBasicParsing `
        -TimeoutSec 30 `
        -ErrorAction Stop

    if ($Respuesta.StatusCode -ne 200) {
        throw "Unexpected HTTP status: $($Respuesta.StatusCode)"
    }

    Write-Host "SPLA INVENTORY SENT SUCCESSFULLY"
    Write-Host "Computer: $env:COMPUTERNAME"
    Write-Host "HTTP: $($Respuesta.StatusCode)"
    exit 0
}
catch {
    Write-Host "SPLA INVENTORY FAILED: $($_.Exception.Message)"
    exit 1
}

