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
    [parameter(Mandatory)]
    [String]$SiteId,
    [parameter(Mandatory)]
    [String]$DriveId,
    [parameter(Mandatory)]
    [String]$FolderId,
    [parameter(Mandatory)]
    [String]$PrinterName,
    [parameter(Mandatory)]
    [Int]$PrinterPort
)

try {
    $ErrorActionPreference = 'Stop'

    $result = @{
        DateTime         = Get-Date
        SiteId           = $SiteId
        DriveId          = $DriveId
        FolderId         = $FolderId
        Printer          = $PrinterName
        PrinterPort      = $PrinterPort
        FileName         = $null
        FileCreationDate = $null
        Actions          = @()
        Error            = $null
        Printed          = $false
    }

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

    Write-Verbose "Found most recent file '$($mostRecentFile.Name)' craated '$($mostRecentFile.CreatedDateTime)'"

    if (-not $mostRecentFile) {
        throw 'No files in folder'
    }

    $result.FileCreationDate = $mostRecentFile.CreatedDateTime
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

    $result.Actions += "File downloaded to '$($downloadFilePath)'"
    #endregion

    #region Print document
    Write-Verbose "Print file on printer '$PrinterName'"

    $params = @{
        FilePath    = $downloadFilePath
        PrinterName = $PrinterName
        PrinterPort = $PrinterPort
    }
    Out-PrintFileHC @params

    $result.Actions += "Printed downloaded file"
    #endregion

    $result.Printed = $true
}
catch {
    Write-Warning "Failed: $_"
    $result.Error = $_
}
finally {
    $result
}