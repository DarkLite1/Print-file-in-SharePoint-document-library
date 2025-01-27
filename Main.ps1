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

            #region Test integer value
            try {
                [int]$MaxConcurrentJobs = $jsonFileContent.MaxConcurrentJobs
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$($jsonFileContent.MaxConcurrentJobs)' is not supported."
            }
            #endregion

            $Tasks = $jsonFileContent.Tasks

            foreach ($task in $Tasks) {
                @(
                    'TaskName', 'Sftp', 'Actions', 'Option'
                ).where(
                    { -not $task.$_ }
                ).foreach(
                    { throw "Property 'Tasks.$_' not found" }
                )

                if (-not $task.TaskName) {
                    throw "Property 'Tasks.TaskName' not found"
                }

                @('ComputerName', 'Credential').where(
                    { -not $task.Sftp.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Sftp.$_' not found" }
                )

                @('UserName').Where(
                    { -not $task.Sftp.Credential.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Sftp.Credential.$_' not found" }
                )

                if (
                    $task.Sftp.Credential.Password -and
                    $task.Sftp.Credential.PasswordKeyFile
                ) {
                    throw "Property 'Tasks.Sftp.Credential.Password' and 'Tasks.Sftp.Credential.PasswordKeyFile' cannot be used at the same time"
                }

                if (
                    (-not $task.Sftp.Credential.Password) -and
                    (-not $task.Sftp.Credential.PasswordKeyFile)
                ) {
                    throw "Property 'Tasks.Sftp.Credential.Password' or 'Tasks.Sftp.Credential.PasswordKeyFile' not found"
                }

                #region Test boolean values
                foreach (
                    $boolean in
                    @(
                        'OverwriteFile'
                    )
                ) {
                    try {
                        $null = [Boolean]::Parse($task.Option.$boolean)
                    }
                    catch {
                        throw "Property 'Tasks.Option.$boolean' is not a boolean value"
                    }
                }
                #endregion

                #region Test file extensions
                $task.Option.FileExtensions.Where(
                    { $_ -and ($_ -notLike '.*') }
                ).foreach(
                    { throw "Property 'Tasks.Option.FileExtensions' needs to start with a dot. For example: '.txt', '.xml', ..." }
                )
                #endregion

                if (-not $task.Actions) {
                    throw 'Tasks.Actions is missing'
                }

                #region Test unique ComputerName
                $task.Actions | Group-Object -Property {
                    $_.ComputerName
                } |
                Where-Object { $_.Count -ge 2 } | ForEach-Object {
                    throw "Duplicate 'Tasks.Actions.ComputerName' found: $($_.Name)"
                }
                #endregion

                foreach ($action in $task.Actions) {
                    if ($action.PSObject.Properties.Name -notContains 'ComputerName') {
                        throw "Property 'Tasks.Actions.ComputerName' not found"
                    }

                    @('Paths').Where(
                        { -not $action.$_ }
                    ).foreach(
                        { throw "Property 'Tasks.Actions.$_' not found" }
                    )

                    foreach ($path in $action.Paths) {
                        @(
                            'Source', 'Destination'
                        ).Where(
                            { -not $path.$_ }
                        ).foreach(
                            {
                                throw "Property 'Tasks.Actions.Paths.$_' not found"
                            }
                        )

                        if (
                            (
                                ($path.Source -like '*/*') -and
                                ($path.Destination -like '*/*')
                            ) -or
                            (
                                ($path.Source -like '*\*') -and
                                ($path.Destination -like '*\*')
                            ) -or
                            (
                                ($path.Source -like 'sftp*') -and
                                ($path.Destination -like 'sftp*')
                            ) -or
                            (
                                -not (
                                    ($path.Source -like 'sftp:/*') -or
                                    ($path.Destination -like 'sftp:/*')
                                )
                            )
                        ) {
                            throw "Property 'Tasks.Actions.Paths.Source' and 'Tasks.Actions.Paths.Destination' needs to have one SFTP path ('sftp:/....') and one folder path (c:\... or \\server$\...). Incorrect values: Source '$($path.Source)' Destination '$($path.Destination)'"
                        }
                    }

                    #region Test unique Source Destination
                    $action.Paths | Group-Object -Property 'Source' |
                    Where-Object { $_.Count -ge 2 } | ForEach-Object {
                        throw "Duplicate 'Tasks.Actions.Paths.Source' found: '$($_.Name)'. Use separate Tasks to run them sequentially instead of in Actions, which is ran in parallel"
                    }
                    #endregion

                    #region Test unique Source Destination
                    $action.Paths | Group-Object -Property 'Destination' |
                    Where-Object { $_.Count -ge 2 } | ForEach-Object {
                        throw "Duplicate 'Tasks.Actions.Paths.Destination' found: '$($_.Name)'. Use separate Tasks to run them sequentially instead of in Actions, which is ran in parallel"
                    }
                    #endregion
                }
            }

            #region Test unique TaskName
            $Tasks.TaskName | Group-Object | Where-Object {
                $_.Count -gt 1
            } | ForEach-Object {
                throw "Property 'Tasks.TaskName' with value '$($_.Name)' is not unique. Each task name needs to be unique."
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
        #region Create general output folder
        $outputFolder = Join-Path -Path $DropFolder -ChildPath 'Output'

        $null = New-Item -Path $outputFolder -ItemType Directory -EA Ignore
        #endregion

        $inputFiles = @()

        foreach ($file in $dropFolderExcelFiles) {
            try {
                $inputFile = [PSCustomObject]@{
                    ExcelFile = @{
                        Item         = $file
                        Content      = @()
                        OutputFolder = $null
                        Error        = $null
                    }
                    FilePath  = @{
                        DownloadResults = $null
                    }
                    Tasks     = @()
                    Error     = $null
                }

                #region Test if file is still present
                if (-not (Test-Path -LiteralPath $inputFile.ExcelFile.Item.FullName -PathType 'Leaf')) {
                    throw "Excel file '$($inputFile.ExcelFile.Item.FullName)' was removed during execution"
                }
                #endregion

                #region Create Excel specific output folder
                try {
                    $params = @{
                        Path        = '{0}\{1} {2}' -f
                        $outputFolder, $startDate, $inputFile.ExcelFile.Item.BaseName
                        ItemType    = 'Directory'
                        Force       = $true
                        ErrorAction = 'Stop'
                    }
                    $inputFile.ExcelFile.OutputFolder = (New-Item @params).FullName

                    Write-Verbose "Excel file output folder '$($inputFile.ExcelFile.OutputFolder)'"

                    $inputFile.FilePath.DownloadResults = Join-Path $inputFile.ExcelFile.OutputFolder 'Download results.xlsx'
                }
                Catch {
                    throw "Failed creating the Excel output folder '$($inputFile.ExcelFile.OutputFolder)': $_"
                }
                #endregion

                try {
                    #region Move original Excel file to output folder
                    try {
                        $moveParams = @{
                            LiteralPath = $inputFile.ExcelFile.Item.FullName
                            Destination = Join-Path $inputFile.ExcelFile.OutputFolder $inputFile.ExcelFile.Item.Name
                            ErrorAction = 'Stop'
                        }

                        Write-Verbose "Move original Excel file '$($moveParams.LiteralPath)' to output folder '$($moveParams.Destination)'"

                        Move-Item @moveParams
                    }
                    catch {
                        $M = $_
                        $error.RemoveAt(0)
                        throw "Failed moving the file '$($inputFile.ExcelFile.Item.FullName)' to folder '$($inputFile.ExcelFile.OutputFolder)': $M"
                    }
                    #endregion

                    #region Import Excel file
                    try {
                        $M = "Import Excel file '$($inputFile.ExcelFile.Item.FullName)'"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                        $params = @{
                            Path          = $moveParams.Destination
                            WorksheetName = $ExcelFileWorksheetName
                            ErrorAction   = 'Stop'
                            DataOnly      = $true
                        }
                        $inputFile.ExcelFile.Content += Import-Excel @params |
                        Select-Object -Property 'Url', 'FileName',
                        'DownloadFolderName'

                        $M = "Imported {0} rows from Excel file '{1}'" -f
                        $inputFile.ExcelFile.Content.count, $inputFile.ExcelFile.Item.FullName
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                    }
                    catch {
                        $error.RemoveAt(0)
                        throw "Worksheet '$($params.WorksheetName)' not found"
                    }
                    #endregion

                    #region Test Excel file
                    foreach ($row in $inputFile.ExcelFile.Content) {
                        if (-not ($row.FileName)) {
                            throw "Property 'FileName' not found"
                        }
                        if (-not ($row.URL)) {
                            throw "Property 'URL' not found"
                        }
                        if (-not ($row.DownloadFolderName)) {
                            throw "Property 'DownloadFolderName' not found"
                        }
                    }
                    #endregion
                }
                catch {
                    Write-Warning "Excel input file error: $_"
                    $inputFile.ExcelFile.Error = $_

                    #region Create Error.html file
                    "
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <style>
                    .myDiv {
                    border: 5px outset red;
                    background-color: lightblue;
                    text-align: center;
                    }
                    </style>
                    </head>
                    <body>

                    <h1>Error detected in the Excel sheet</h1>

                    <div class=`"myDiv`">
                    <h2>$_</h2>
                    </div>

                    <p>Please fix this error and try again.</p>

                    </body>
                    </html>
                    " | Out-File -LiteralPath "$($inputFile.ExcelFile.OutputFolder)\Error.html" -Encoding utf8
                    #endregion

                    $error.RemoveAt(0)
                    Continue
                }

                $progressCount = @{
                    Current = 0
                    Total   = $inputFile.ExcelFile.Content.count
                }

                foreach (
                    $collection in
                    ($inputFile.ExcelFile.Content |
                    Group-Object -Property 'DownloadFolderName')
                ) {
                    try {
                        $task = [PSCustomObject]@{
                            ItemsToDownload = $collection.Group
                            DownloadFolder  = @{
                                Name = $collection.Name
                                Path = Join-Path $inputFile.ExcelFile.OutputFolder "Downloads\Files\$($collection.Name)"
                            }
                            Job             = @{
                                Object = @()
                                Result = @()
                            }
                            FilePath        = @{
                                ZipFile = Join-Path $inputFile.ExcelFile.OutputFolder "Downloads\Zip files\$($collection.Name).zip"
                            }
                            Error           = $null
                        }

                        #region Create download folder
                        try {
                            $params = @{
                                Path        = $task.DownloadFolder.Path
                                ItemType    = 'Directory'
                                ErrorAction = 'Stop'
                            }
                            $task.DownloadFolder.Item = New-Item @params
                        }
                        catch {
                            throw "Failed creating download folder '$($params.Path)': $_"
                        }
                        #endregion

                        #region Download files
                        $M = "Download $($task.ItemsToDownload.count) files to '$($task.DownloadFolder.Path)'"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                        foreach ($row in $task.ItemsToDownload) {
                            $progressCount.Current++

                            $M = "Download {0}/{1} file name '$($row.FileName)' from '$($row.Url)'" -f $progressCount.Current, $progressCount.Total
                            Write-Verbose $M

                            $task.Job.Object += Start-Job -ScriptBlock {
                                Param (
                                    [Parameter(Mandatory)]
                                    [String]$Url,
                                    [Parameter(Mandatory)]
                                    [String]$DownloadFolder,
                                    [Parameter(Mandatory)]
                                    [String]$FileName
                                )

                                try {
                                    $result = [PSCustomObject]@{
                                        Url                = $Url
                                        FileName           = $FileName
                                        FilePath           = $null
                                        DownloadFolderName = Split-Path $DownloadFolder -Leaf
                                        DownloadedOn       = $null
                                        Error              = $null
                                    }

                                    $result.FilePath = Join-Path -Path $DownloadFolder -ChildPath $FileName

                                    $invokeParams = @{
                                        Uri         = $result.Url
                                        OutFile     = $result.FilePath
                                        TimeoutSec  = 10
                                        ErrorAction = 'Stop'
                                    }
                                    $null = Invoke-WebRequest @invokeParams

                                    $result.DownloadedOn = Get-Date
                                }
                                catch {
                                    $statusCode = $_.Exception.Response.StatusCode.value__

                                    if ($statusCode) {
                                        $errorMessage = switch ($statusCode) {
                                            '404' {
                                                'Status code: 404 Not found'; break
                                            }
                                            Default {
                                                "Status code: $_"
                                            }
                                        }
                                    }
                                    else {
                                        $errorMessage = $_
                                    }

                                    $result.Error = "Download failed: $errorMessage"
                                    $Error.RemoveAt(0)
                                }
                                finally {
                                    $result
                                }
                            } -ArgumentList $row.Url, $task.DownloadFolder.Path, $row.FileName

                            #region Wait for max running jobs
                            $waitParams = @{
                                Name       = $task.Job.Object | Where-Object { $_ }
                                MaxThreads = $MaxConcurrentJobs
                            }
                            Wait-MaxRunningJobsHC @waitParams
                            #endregion
                        }
                        #endregion

                        #region Wait for jobs to finish
                        $M = "Wait for all $($task.Job.Object.count) jobs to finish"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                        $null = $task.Job.Object | Wait-Job
                        #endregion

                        #region Get job results and job errors
                        $task.Job.Result += $task.Job.Object | Receive-Job
                        #endregion

                        #region Create zip file
                        if (
                            ($task.ItemsToDownload.Count) -eq
                            ($task.Job.Result.where({ $_.DownloadedOn }).count)
                        ) {
                            try {
                                $M = "Create zip file with $($task.Job.Result.count) files in zip file '$($task.FilePath.ZipFile)'"
                                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                                $Source = $task.DownloadFolder.Path
                                $Target = $task.FilePath.ZipFile
                                Start-SevenZip a -mx=9 $Target $Source

                                if ($LASTEXITCODE -ne 0) {
                                    throw "7 zip failed with last exit code: $LASTEXITCODE"
                                }
                            }
                            catch {
                                $M = $_
                                $Error.RemoveAt(0)
                                throw "Failed creating zip file: $M"
                            }
                        }
                        else {
                            $M = 'Not all files downloaded, no zip file created'
                            Write-Verbose $M; Write-EventLog @EventWarnParams -Message $M

                            #region Create Error.html file
                            "
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <style>
                    .myDiv {
                    border: 5px outset red;
                    background-color: lightblue;
                    text-align: center;
                    }
                    </style>
                    </head>
                    <body>

                    <h1>Error detected while downloading files</h1>

                    <div class=`"myDiv`">
                    <h1>DownloadFolderName '$($task.DownloadFolder.Name)'</h1>
                    <h2>No zip-file created because not all files could be downloaded.</h2>
                    </div>

                    <h1>Detected errors:</h>
                    <ul>
                        $($task.Job.Result.where({ $_.Error }).foreach(
                            { '<li>{0} - {1}</li>' -f $_.FileName, $_.Error }
                            )
                        )
                    </ul>

                    <p>Please check the Excel file '$($inputFile.FilePath.DownloadResults)' for more information.</p>

                    </body>
                    </html>
                    " | Out-File -LiteralPath "$($inputFile.ExcelFile.OutputFolder)\Error - $($task.DownloadFolder.Name).html" -Encoding utf8
                            #endregion
                        }
                        #endregion
                    }
                    catch {
                        $M = $_
                        Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M

                        $task.Error = $_
                        $error.RemoveAt(0)
                    }
                    finally {
                        $inputFile.Tasks += $task
                    }
                }

                #region Export results to Excel
                if ($inputFile.Tasks.Job.Result) {
                    $excelParams = @{
                        Path               = $inputFile.FilePath.DownloadResults
                        NoNumberConversion = '*'
                        WorksheetName      = 'Overview'
                        TableName          = 'Overview'
                        AutoSize           = $true
                        FreezeTopRow       = $true
                    }

                    $M = "Export $($inputFile.Tasks.Job.Result.Count) rows to Excel file '$($excelParams.Path)'"
                    Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                    $inputFile.Tasks.Job.Result |
                    Select-Object -Property 'Url',
                    'FileName', 'DownloadFolderName',
                    'DownloadedOn' , 'Error', 'FilePath' |
                    Export-Excel @excelParams
                }
                #endregion
            }
            catch {
                $M = $_
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M

                $inputFile.Error = $_
                $error.RemoveAt(0)
            }
            finally {
                $inputFiles += $inputFile
            }
        }
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