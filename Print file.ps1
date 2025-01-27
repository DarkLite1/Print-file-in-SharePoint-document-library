#Requires -Version 7
#Requires -Modules Toolbox.FileAndFolder

<#
    .SYNOPSIS
        Print a file in the SharePoint document library

    .PARAMETER SiteId
        ID of the SharePoint site

        $site = Get-MgSite -SiteId 'hcgroupnet.sharepoint.com:/sites/BEL_xx'
        $site.id

    .PARAMETER DriveId
        ID of the SharePoint drive

        $drives = Get-MgSiteDrive -SiteId $site.Id
        $drive.id

    .PARAMETER FolderId
        ID of the SharePoint folder

        $driveRootChildren = Get-MgDriveRootChild -DriveId $drive.id
        $driveRootChildren.id

        Get-MgDriveItem -DriveId $DriveId -DriveItemId $driveRootChildren.id
#>

[CmdLetBinding()]
Param (
    [String]$SiteId = 'hcgroupnet.sharepoint.com,b4c482ba-d46d-4a40-93f1-463b40faacd4,213a6ffc-1009-43ca-be81-d20b54789765',
    [String]$DriveId = 'b!uoLEtG3UQEqT8UY7QPqs1PxvOiEJEMpDvoHSC1R4l2W5o0sx337DTZPZGtpfnBvg',
    [String]$FolderId = '01P4GU2YRPPOUAAVJA6ZC2IZMNPRNCVPFI',
    [String]$PrinterName = 'BELPRLIXH609',
    [Int]$PrinterPort = 9100

)

Start-Transcript -Path "T:\Test\Brecht\PowerShell\Print file in SharePoint document library\Transcript.txt" -Append -UseMinimalHeader

$ErrorActionPreference = 'Stop'
$VerbosePreference = 'Continue'

#region Connect to Microsoft Graph
Write-Verbose 'Connect to MS Graph'

$params = @{
    ClientId              = $env:AZURE_CLIENT_ID
    TenantId              = $env:AZURE_TENANT_ID
    CertificateThumbprint = $env:AZURE_POWERSHELL_CERTIFICATE_THUMBPRINT
    NoWelcome             = $true
}
Connect-MgGraph @params
#endregion

#region Test SiteId
try {
    Write-Verbose "Test site ID '$SiteId' exists"

    $site = Get-MgSite -SiteId $SiteId -ErrorAction Stop
}
catch {
    throw "Site id '$SiteId' not found: $_"
}
#endregion

#region Test DriveId
try {
    Write-Verbose "Test drive ID '$DriveId' exists"

    $drives = Get-MgSiteDrive -SiteId $SiteId

    if ($drives.Id -notContains $DriveId) {
        throw "Drive id '$DriveId' not found"
    }
}
catch {
    throw "Failed retrieving site drive: $_"
}
#endregion

#region Test FolderId
Write-Verbose "Test folder ID '$FolderId' exists"

$driveRootChildren = Get-MgDriveRootChild -DriveId $DriveId

if ($driveRootChildren.Id -notContains $FolderId) {
    throw "FolderId '$FolderId' not found"
}

Write-Verbose "Test folder ID '$FolderId' is a folder"

$driveItem = Get-MgDriveItem -DriveId $DriveId -DriveItemId $FolderId

if ($driveItem.Folder -eq $null) {
    throw "FolderId '$FolderId' is not a folder"
}
#endregion

#region Get files in folder
Write-Verbose 'Get files in folder'

$driveItemChild = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $FolderId

if (-not $driveItemChild) {
    throw "No files in folder '$FolderId'"
}
#endregion

#region Get latest file
Write-Verbose 'Select latest file'

$mostRecentFile = $driveItemChild | Sort-Object -Property CreatedDateTime -Descending | Select-Object -First 1

Write-Verbose "Most recent file '$($mostRecentFile.Name)'"
#endregion

#region Download file
$downloadFilePath = '{0}\{1}' -f $env:TEMP, $mostRecentFile.Name

$params = @{
    DriveId     = $DriveId
    DriveItemId = $mostRecentFile.Id
    OutFile     = $downloadFilePath
}

Write-Verbose "Download file to '$($params.OutFile)'"

Get-MgDriveItemContent @params
#endregion

#region Print document
Write-Verbose "Print file on printer '$PrinterName'"

$params = @{
    FilePath                     = $downloadFilePath
    PrinterName                  = $PrinterName
    PrinterPort                  = $PrinterPort
}
Out-PrintFileHC @params
#endregion

Stop-Transcript