#Requires -Version 7
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
            [Int]$PrinterPort
        )

        $params = @{
            SiteId      = $SiteId
            DriveId     = $DriveId
            FolderId    = $FolderId
            PrinterName = $PrinterName
            PrinterPort = $PrinterPort
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
                'SharePoint', 'Printer', 'SendMail', 'ExportExcelFile'
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

            #region Test ExportExcelFile
            @('When').Where(
                { -not $jsonFileContent.ExportExcelFile.$_ }
            ).foreach(
                { throw "Property 'ExportExcelFile.$_' not found" }
            )

            if ($jsonFileContent.ExportExcelFile.When -notMatch '^Never$|^OnlyOnError$|^OnlyOnErrorOrAction$') {
                throw "Property 'ExportExcelFile.When' with value '$($jsonFileContent.ExportExcelFile.When)' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'"
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
        #region Run Print file script
        $M = "Run Print file script"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $params = @{
            SiteId      = $jsonFileContent.SharePoint.SiteId
            DriveId     = $jsonFileContent.SharePoint.DriveId
            FolderId    = $jsonFileContent.SharePoint.FolderId
            PrinterName = $jsonFileContent.Printer.Name
            PrinterPort = $jsonFileContent.Printer.Port
        }
        $result = Invoke-PrintFileScriptHC @params
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
        if (-not $inputFiles) {
            Write-Verbose "No tasks found, exit script"
            Write-EventLog @EventEndParams; Exit
        }

        # $M = "Wait for all $($inputFile.Job.Object.count) jobs to finish"
        # Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $mailParams = @{ }
        $htmlTableTasks = @()

        #region Count totals
        $totalCounter = @{
            All          = @{
                Errors          = 0
                RowsInExcel     = 0
                DownloadedFiles = 0
            }
            SystemErrors = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }

        $totalCounter.All.Errors += $totalCounter.SystemErrors
        #endregion

        foreach ($inputFile in $inputFiles) {
            #region Count task results
            $counter = @{
                RowsInExcel     = (
                    $inputFile.ExcelFile.Content | Measure-Object
                ).Count
                DownloadedFiles = (
                    $inputFile.Tasks.Job.Result.Where({ $_.DownloadedOn }) | Measure-Object
                ).Count
                Errors          = @{
                    InExcelFile      = (
                        $inputFile.ExcelFile.Error | Measure-Object
                    ).Count
                    InTasks          = (
                        $inputFile.Tasks.Where({ $_.Error }) | Measure-Object
                    ).Count
                    DownloadingFiles = (
                        $inputFile.Tasks.Job.Result.Where({ $_.Error }) | Measure-Object
                    ).Count
                    Other            = (
                        $inputFile.Error | Measure-Object
                    ).Count
                }
            }

            $totalCounter.All.RowsInExcel += $counter.RowsInExcel
            $totalCounter.All.DownloadedFiles += $counter.DownloadedFiles
            $totalCounter.All.Errors += (
                $counter.Errors.InExcelFile +
                $counter.Errors.InTasks +
                $counter.Errors.DownloadingFiles +
                $counter.Errors.Other
            )
            #endregion

            #region Create HTML table
            $htmlTableTasks += "
                <table>
                <tr>
                    <th colspan=`"2`">$($inputFile.ExcelFile.Item.Name)</th>
                </tr>
                <tr>
                    <td>Details</td>
                    <td>
                        <a href=`"$($inputFile.ExcelFile.OutputFolder)`">Output folder</a>
                    </td>
                </tr>
                <tr>
                    <td>$($counter.RowsInExcel)</td>
                    <td>Files to download</td>
                </tr>
                <tr>
                    <td>$($counter.DownloadedFiles)</td>
                    <td>Files successfully downloaded</td>
                </tr>
                $(
                    if ($counter.Errors.InExcelFile) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.InExcelFile)</td>
                            <td style=``"background-color: red``">Error{0} in the Excel file</td>
                        </tr>" -f $(if ($counter.Errors.InExcelFile -ne 1) {'s'})
                    }
                )
                $(
                    if ($counter.Errors.DownloadingFiles) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.DownloadingFiles)</td>
                            <td style=``"background-color: red``">File{0} failed to download</td>
                        </tr>" -f $(if ($counter.Errors.DownloadingFiles -ne 1) {'s'})
                    }
                )
                $(
                    if ($counter.Errors.Other) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.Other)</td>
                            <td style=``"background-color: red``">Error{0} found:<br>{1}</td>
                        </tr>" -f $(
                            if ($counter.Errors.Other -ne 1) {'s'}
                        ),
                        (
                            '- ' + $($inputFile.Error -join '<br> - ')
                        )
                    }
                )
                $(
                    if($inputFile.Tasks) {
                        "<tr>
                            <th colspan=``"2``">Downloads per folder</th>
                        </tr>"
                    }
                )
                $(
                    foreach (
                        $task in
                        (
                            $inputFile.Tasks |
                            Sort-Object {$_.DownloadFolder.Name}
                        )
                    ) {
                        $errorCount = $task.Job.Result.Where(
                            {$_.Error}).Count

                        $template = if ($errorCount) {
                            "<tr>
                            <td style=``"background-color: red``">{0}/{1}</td>
                            <td style=``"background-color: red``">{2}{3}</td>
                            </tr>"
                        } else {
                            "<tr>
                                <td>{0}/{1}</td>
                                <td>{2}{3}</td>
                            </tr>"
                        }

                        $template -f
                        $(
                            $task.Job.Result.Where({$_.DownloadedOn}).Count
                        ),
                        $(
                            ($task.ItemsToDownload | Measure-Object).Count
                        ),
                        $(
                            $task.DownloadFolder.Name
                        ),
                        $(
                            if ($errorCount) {
                                ' ({0} error{1})' -f
                                $errorCount, $(if ($errorCount -ne 1) {'s'})
                            }
                        )
                    }
                )
            </table>
            "
            #endregion
        }

        $htmlTableTasks = $htmlTableTasks -join '<br>'

        #region Send summary mail to user

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} file{2} downloaded' -f
        $totalCounter.All.DownloadedFiles,
        $totalCounter.All.RowsInExcel,
        $(
            if ($totalCounter.All.RowsInExcel -ne 1) {
                's'
            }
        )

        if (
            $totalErrorCount = $totalCounter.All.Errors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create error html lists
        $systemErrorsHtmlList = if ($totalCounter.SystemErrors) {
            "<p>Detected <b>{0} system error{1}</b>:{2}</p>" -f $totalCounter.SystemErrors,
            $(
                if ($totalCounter.SystemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } |
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $htmlTableTasks"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
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