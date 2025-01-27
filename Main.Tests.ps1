﻿#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        SharePoint      = @{
            SiteId   = 'theSiteId'
            DriveId  = 'theDriveId'
            FolderId = 'theFolderId'
        }
        Printer = @{
            Name = 'thePrinterName'
            Port = 9100
        }
        SendMail        = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        ExportExcelFile = @{
            When = 'OnlyOnErrorOrAction'
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        ScriptPath  = @{
            PrintFile = (New-Item 'TestDrive:/u.ps1' -ItemType 'File').FullName
        }
        LogFolder   = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the file is not found' {
        It 'ScriptPath.PrintFile' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.ScriptPath.PrintFile = 'c:\upDoesNotExist.ps1'

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*ScriptPath.PrintFile 'c:\upDoesNotExist.ps1' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find Path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'SharePoint', 'Printer', 'SendMail', 'ExportExcelFile'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.<_> not found' -ForEach @(
                'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'ExportExcelFile.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.ExportExcelFile.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'ExportExcelFile.When' with value 'wrong' is not valid. Accepted values are 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SendMail.When is not valid' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.When = 'wrong'

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.When' with value 'wrong' is not valid. Accepted values are 'Always', 'Never', 'OnlyOnError' or 'OnlyOnErrorOrAction'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'correct the import file' {
    Context "add trailing slashes to Paths starting with 'sftp:/'" {
        It 'Source' {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].Actions[0].Paths[0].Source = 'sftp:/a'
            $testNewInputFile.Tasks[0].Actions[0].Paths[0].Destination = 'TestDrive:\b'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            $Tasks[0].Actions[0].Paths[0].Source | Should -Be 'sftp:/a/'
            $Tasks[0].Actions[0].Paths[0].Destination | Should -Be 'TestDrive:\b'
        }
        It 'Destination' {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].Actions[0].Paths[0].Source = 'TestDrive:\b'
            $testNewInputFile.Tasks[0].Actions[0].Paths[0].Destination = 'sftp:/a'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            $Tasks[0].Actions[0].Paths[0].Source | Should -Be 'TestDrive:\b'
            $Tasks[0].Actions[0].Paths[0].Destination | Should -Be 'sftp:/a/'
        }
    }
}
Describe 'execute the SFTP script when' {
    BeforeAll {
        $testJobArguments = @(
            {
                ($Session) -and
                ($FilePath -eq $testParams.ScriptPath.PrintFile) -and
                ($ArgumentList[0] -eq $testInputFile.Tasks[0].Sftp.ComputerName) -and
                ($ArgumentList[1].GetType().Name -eq 'PSCredential') -and
                ($ArgumentList[2].GetType().BaseType.Name -eq 'Array') -and
                ($ArgumentList[3] -eq $testInputFile.MaxConcurrentJobs) -and
                (-not $ArgumentList[4]) -and
                ($ArgumentList[5] -eq $testInputFile.Tasks[0].Option.FileExtensions) -and
                ($ArgumentList[6] -eq $testInputFile.Tasks[0].Option.OverwriteFile)
            }
        )

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Actions = @(
            $testNewInputFile.Tasks[0].Actions[0]
        )
    }
    Context 'Tasks.Actions.ComputerName is not the localhost' {
        BeforeAll {
            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'call New-PSSession' {
            Should -Invoke New-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ErrorAction -eq 'SilentlyContinue') -and
                ($ConfigurationName -eq $PSSessionConfiguration) -and
                ($ComputerName -eq $testNewInputFile.Tasks[0].Actions[0].ComputerName)
            }
        }
        It 'call Invoke-Command' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                (& $testJobArguments[0])
            }
        }
        It 'call Remove-PSSession' {
            Should -Invoke Remove-PSSession -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($Session -eq $testPsSession)
            }
        }
    }
    Context 'Tasks.Actions.ComputerName is the localhost' {
        BeforeAll {
            $testNewInputFile.Tasks[0].Actions[0].ComputerName = 'localhost'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'do not call Invoke-Command ' {
            Should -Not -Invoke Invoke-Command -Scope Context
        }
        It 'do not call New-PSSession ' {
            Should -Not -Invoke New-PSSession -Scope Context
        }
        It 'do not call Remove-PSSession ' {
            Should -Not -Invoke Remove-PSSession -Scope Context
        }
    }
    Context 'with Tasks.Sftp.Credential.PasswordKeyFile' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks[0].Actions = @(
                $testNewInputFile.Tasks[0].Actions[0]
            )

            'sadfsd' | Out-File 'TestDrive:\key.txt'

            $testNewInputFile.Tasks[0].Sftp.Credential.Password = $null
            $testNewInputFile.Tasks[0].Sftp.Credential.PasswordKeyFile = 'TestDrive:\key.txt'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        It 'call Invoke-Command' {
            Should -Invoke Invoke-Command -Times 1 -Exactly -Scope Context -ParameterFilter {
                ($ArgumentList[4] -ne $null)
            }
        }
    }
}
Describe 'when the SFTP script runs successfully' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams
    }
    Context 'create an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter "* - $((Split-Path $testOutParams.FilePath -Leaf).TrimEnd('.json')) - Log.xlsx"

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'in the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Source -eq $testRow.Source
                }
                $actualRow.TaskName | Should -Be $testRow.TaskName
                $actualRow.SftpServer | Should -Be $testRow.SftpServer
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.Destination | Should -Be $testRow.Destination
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.Action | Should -Be $testRow.Action
                $actualRow.FileName | Should -Be $testRow.FileName
                $actualRow.FileSize | Should -Be $testRow.FileSize
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.SendMail.To) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2 moved') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*Summary of SFTP actions*table*$($testInputFile.Tasks[0].TaskName)*$($testInputFile.Tasks[0].Sftp.ComputerName)*Source*Destination*Result*$($testInputFile.Tasks[0].Actions[0].Paths[0].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[0].Destination)*1 moved*$($testInputFile.Tasks[0].Actions[0].Paths[1].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[1].Destination)*1 moved*2 moved on $($testInputFile.Tasks[0].Actions[0].ComputerName)*")
            }
        }
    }
}
Describe 'ExportExcelFile.When' {
    Context 'create no Excel file' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'Never'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Invoke-Command {
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
    }
    Context 'create an Excel file' {
        It "'OnlyOnError' and there are errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Uploaded = $true
                    Action   = @('upload')
                    Error    = $null
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    Uploaded = $false
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.ExportExcelFile.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
    }
}
Describe 'SendMail.When' {
    BeforeAll {
        $testParamFilter = @{
            ParameterFilter = { $To -eq $testNewInputFile.SendMail.To }
        }
    }
    Context 'send no e-mail to the user' {
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'Never'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Invoke-Command {
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
    }
    Context 'send an e-mail to the user' {
        It "'OnlyOnError' and there are errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Uploaded = $true
                    Action   = @('upload')
                    Error    = $null
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command {
                [PSCustomObject]@{
                    Path     = 'a'
                    DateTime = Get-Date
                    Action   = @()
                    Error    = 'oops'
                }
            } -ParameterFilter {
                $FilePath -eq $testParams.ScriptPath.PrintFile
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
    }
}
Describe 'ReportOnly' {
    BeforeAll {
    }
    Context 'when no previously exported Excel file is found' {
        BeforeAll {
            Get-ChildItem $testParams.LogFolder -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams -ReportOnly
        }
        It 'no not create an Excel file' {
            Get-ChildItem $testParams.LogFolder -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
        It 'do not call the SFTP script' {
            Should -Not -Invoke New-PSSession
            Should -Not -Invoke Invoke-Command
        }
        It 'send an e-mail' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context

            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $testInputFile.SendMail.To) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '0 moved') -and
                (-not $Attachments) -and
                ($Message -like "*Summary of all SFTP actions <b>executed today</b>*table*$($testInputFile.Tasks[0].TaskName)*$($testInputFile.Tasks[0].Sftp.ComputerName)*Source*Destination*Result*$($testInputFile.Tasks[0].Actions[0].Paths[0].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[0].Destination)*0 moved*$($testInputFile.Tasks[0].Actions[0].Paths[1].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[1].Destination)*0 moved*0 moved on $($testInputFile.Tasks[0].Actions[0].ComputerName)*")
            }
        }
    }
    Context 'when a previously exported Excel file is found' {
        BeforeAll {
            $testExportParams = @{
                WorksheetName = 'Overview'
                Path          = $testParams.LogFolder + '\' + (Get-Date).ToString('yyyy-MM-dd') + ' - ' + $testParams.ScriptName + ' - ' + (Split-Path $testParams.ImportFile -Leaf).TrimEnd('.json') + ' - Log.xlsx'
            }
            $testExportedExcelRows | Export-Excel @testExportParams

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams -ReportOnly
        }
        It 'do not call the SFTP script' {
            Should -Not -Invoke New-PSSession
            Should -Not -Invoke Invoke-Command
        }
        It 'send an e-mail' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq $testInputFile.SendMail.To) -and
            ($Attachments -eq $testExportParams.Path) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2 moved') -and
            ($Message -like "*Summary of all SFTP actions <b>executed today</b>*table*$($testInputFile.Tasks[0].TaskName)*$($testInputFile.Tasks[0].Sftp.ComputerName)*Source*Destination*Result*$($testInputFile.Tasks[0].Actions[0].Paths[0].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[0].Destination)*1 moved*$($testInputFile.Tasks[0].Actions[0].Paths[1].Source)*$($testInputFile.Tasks[0].Actions[0].Paths[1].Destination)*1 moved*2 moved on $($testInputFile.Tasks[0].Actions[0].ComputerName)*")
            }
        }
    }
}
