#Requires -Version 7
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.FileAndFolder

<#
    .SYNOPSIS
        Print a file in the SharePoint document library

    .DESCRIPTION
        This script is intended to be triggered by a scheduled task to
        periodically print the latest file found in a SharePoint document
        library.

    .PARAMETER ImportFile
        Contains all the parameters for the script

    .PARAMETER SharePoint.SiteId
        ID of the SharePoint site

        $site = Get-MgSite -SiteId 'hcgroupnet.sharepoint.com:/sites/BEL_xx'
        $site.id

    .PARAMETER SharePoint.DriveId
        ID of the SharePoint drive

        $drives = Get-MgSiteDrive -SiteId $site.Id
        $drive.id

    .PARAMETER SharePoint.FolderId
        ID of the SharePoint folder

        $driveRootChildren = Get-MgDriveRootChild -DriveId $drive.id
        $driveRootChildren.id

        Get-MgDriveItem -DriveId $DriveId -DriveItemId $driveRootChildren.id

    .PARAMETER Printer.Name
        Name or IP address of the printer to print the file

    .PARAMETER Printer.Port
        Port of the printer to print the file

    .PARAMETER Option.PrintDuplicate
        Print the same file multiple times

        Valid values:
        - $true  : Print the same file multiple times
        - $false : Print the same file only once

    .PARAMETER SendMail
        Contains all the information for sending e-mails.

    .PARAMETER SendMail.To
        Destination e-mail addresses.

    .PARAMETER SendMail.When
        When does the script need to send an e-mail.

        Valid values:
        - Always              : Always sent an e-mail
        - Never               : Never sent an e-mail
        - OnlyOnError         : Only sent an e-mail when errors where detected
        - OnlyOnErrorOrAction : Only sent an e-mail when errors where detected
                                or when a file is printed
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [HashTable]$ScriptPath = @{
        PrintFile = "$PSScriptRoot\Print file.ps1"
    },
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    function Invoke-PrintFileScriptHC {
        param (
            [parameter(Mandatory)]
            [String]$SiteId,
            [parameter(Mandatory)]
            [String]$DriveId,
            [parameter(Mandatory)]
            [String]$FolderId,
            [parameter(Mandatory)]
            [String]$PrinterName,
            [parameter(Mandatory)]
            [Int]$PrinterPort,
            [string]$FileNameLastPrintedFile
        )

        $params = @{
            SiteId      = $SiteId
            DriveId     = $DriveId
            FolderId    = $FolderId
            PrinterName = $PrinterName
            PrinterPort = $PrinterPort
        }

        if ($FileNameLastPrintedFile) {
            $params.FileNameLastPrintedFile = $FileNameLastPrintedFile
        }

        & $scriptPathItem.PrintFile @params
    }

    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $Error.Clear()

        #region Test path exists
        $scriptPathItem = @{}

        $ScriptPath.GetEnumerator().ForEach(
            {
                try {
                    $key = $_.Key
                    $value = $_.Value

                    $params = @{
                        Path        = $value
                        ErrorAction = 'Stop'
                    }
                    $scriptPathItem[$key] = (Get-Item @params).FullName
                }
                catch {
                    throw "ScriptPath.$key '$value' not found"
                }
            }
        )
        #endregion

        #region Create log folder
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            LiteralPath = $ImportFile
            Raw         = $true
            ErrorAction = 'Stop'
            Encoding    = 'UTF8'
        }
        $jsonFileContent = Get-Content @params | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        Write-Verbose 'Test .json file properties'

        try {
            @(
                'SharePoint', 'Printer', 'SendMail'
            ).where(
                { -not $jsonFileContent.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            #region Test SharePoint
            @('SiteId', 'DriveId', 'FolderId').Where(
                { -not $jsonFileContent.SharePoint.$_ }
            ).foreach(
                { throw "Property 'SharePoint.$_' not found" }
            )
            #endregion

            #region Test Printer
            @('Name', 'Port').Where(
                { -not $jsonFileContent.Printer.$_ }
            ).foreach(
                { throw "Property 'Printer.$_' not found" }
            )
            #endregion

            #region Test boolean values
            foreach (
                $boolean in
                @(
                    'PrintDuplicate'
                )
            ) {
                try {
                    $null = [Boolean]::Parse($jsonFileContent.Option.$boolean)
                }
                catch {
                    throw "Property 'Option.$boolean' is not a boolean value"
                }
            }
            #endregion

            #region Test SendMail
            @('To', 'When').Where(
                { -not $jsonFileContent.SendMail.$_ }
            ).foreach(
                { throw "Property 'SendMail.$_' not found" }
            )

            if ($jsonFileContent.SendMail.When -notMatch '^Never$|^Always$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'SendMail.When' with value '$($jsonFileContent.SendMail.When)' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
            }
            #endregion
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
Process {
    Try {
        #region Get name of the last printed file
        $fileNameLastPrintedFile = $null

        if (-not $jsonFileContent.Option.PrintDuplicate) {
            $params = @{
                LiteralPath = $logParams.LogFolder
                File        = $true
                Recurse     = $true
                Filter      = "* - Log.xlsx"
            }
            $latestExcelLogFile = Get-ChildItem @params |
            Sort-Object CreationTime -Descending |
            Select-Object -First 1

            if ($latestExcelLogFile) {
                $params = @{
                    Path          = $latestExcelLogFile.FullName
                    WorksheetName = 'Overview'
                }
                $excelLogFileContent = Import-Excel @params |
                Select-Object -First 1

                if (-not $excelLogFileContent.Error) {
                    $fileNameLastPrintedFile = $excelLogFileContent.FileName
                }
            }
        }
        #endregion


        #region Run Print file script
        $M = "Run Print file script"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            SiteId                  = $jsonFileContent.SharePoint.SiteId
            DriveId                 = $jsonFileContent.SharePoint.DriveId
            FolderId                = $jsonFileContent.SharePoint.FolderId
            PrinterName             = $jsonFileContent.Printer.Name
            PrinterPort             = $jsonFileContent.Printer.Port
            FileNameLastPrintedFile = $fileNameLastPrintedFile
        }
        $results = Invoke-PrintFileScriptHC @params
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
End {
    try {
        #region Counter
        $counter = @{
            FilesPrinted = (
                $results | Where-Object { $_.Printed } | Measure-Object).Count
            Errors       = (
                $results | Where-Object { $_.Error } | Measure-Object).Count
            Actions       = (
                $results | Where-Object { $_.Actions } | Measure-Object).Count
        }
        #endregion

        #region Create system error html lists
        $countSystemErrors = (
            $Error.Exception.Message | Measure-Object
        ).Count

        $systemErrorsHtmlList = if ($countSystemErrors) {
            "<p>Detected <b>{0} system error{1}</b>:{2}</p>" -f $countSystemErrors,
            $(
                if ($countSystemErrors -ne 1) { 's' }
            ),
            $(
                $errorList = $Error.Exception.Message | Where-Object { $_ }
                $errorList | ConvertTo-HtmlListHC

                $errorList.foreach(
                    {
                        $M = "System error: $_"
                        Write-Warning $M
                        Write-EventLog @EventErrorParams -Message $M
                    }
                )
            )
        }

        $counter.Errors += $countSystemErrors
        #endregion

        #region Create HTML table
        Write-Verbose 'Create HTML table'

        $htmlTable = "
        <table>
            <tr>
                <th>Printed files</th>
                <td>$($counter.FilesPrinted)</td>
            </tr>
            $(
                if ($counter.Errors) {
                    "<tr>
                        <th>Errors</th>
                        <td>$($counter.Errors)</td>
                    </tr>"
                }
            )
        </table>
        "
        #endregion

        #region Create Excel objects
        $exportToExcel = $results | Select-Object 'DateTime',
        'FileName', 'FileCreationDate',
        @{
            Name       = 'PrintDuplicate'
            Expression = { $jsonFileContent.Option.PrintDuplicate }
        },
        'Printed',
        @{
            Name       = 'Actions'
            Expression = { $_.Actions -join ', ' }
        },
        @{
            Name       = 'Info'
            Expression = { $_.Info -join ', ' }
        },
        'Error'
        #endregion

        $mailParams = @{}

        #region Create Excel file
        if ($exportToExcel) {
            $excelParams = @{
                Path          = "$logFile - log.xlsx"
                TableName     = 'Overview'
                WorksheetName = 'Overview'
                FreezeTopRow  = $true
                AutoSize      = $true
                Verbose       = $false
            }

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $exportToExcel.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $exportToExcel | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Mail subject and priority
        $mailParams += @{
            Priority = 'Normal'
            Subject  = @("$($counter.FilesPrinted) printed")
        }

        if ($counter.Errors) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += "{0} error{1}" -f
            $counter.Errors,
            $(if ($counter.Errors -ne 1) { 's' })
        }

        $mailParams.Subject = $mailParams.Subject -join ', '
        #endregion

        #region Check to send mail to user
        $sendMailToUser = $false

        if (
            (
                $jsonFileContent.SendMail.When -eq 'Always'
            ) -or
            (
                ($jsonFileContent.SendMail.When -eq 'OnlyOnError') -and
                ($counter.Errors)
            ) -or
            (
                ($jsonFileContent.SendMail.When -eq 'OnlyOnErrorOrAction') -and
                (
                    ($counter.Errors) -or
                    ($counter.FilesPrinted) -or
                    ($counter.Actions)
                )
            )
        ) {
            $sendMailToUser = $true
        }
        #endregion

        #region Send mail
        $mailParams += @{
            To             = $jsonFileContent.SendMail.To
            Message        = "
                $systemErrorsHtmlList
                $(
                    '<p>Summary of print actions:</p>'
                )
                $htmlTable"
            LogFolder      = $LogParams.LogFolder
            Header         = $ScriptName
            EventLogSource = $ScriptName
            Save           = $LogFile + ' - Mail.html'
            ErrorAction    = 'Stop'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message +=
            "<p><i>* Check the attachment for details</i></p>"
        }

        Get-ScriptRuntimeHC -Stop

        if ($sendMailToUser) {
            Write-Verbose 'Send e-mail to the user'

            if ($counter.Errors) {
                $mailParams.Bcc = $ScriptAdmin
            }
            Send-MailHC @mailParams
        }
        else {
            Write-Verbose 'Send no e-mail to the user'

            if ($counter.Errors) {
                Write-Verbose 'Send e-mail to admin only with errors'

                $mailParams.To = $ScriptAdmin
                Send-MailHC @mailParams
            }
        }
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}