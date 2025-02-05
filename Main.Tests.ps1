#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML
#Requires -Version 5.1

BeforeAll {
    $testInputFile = @{
        Option          = @{
            PrintDuplicate = $true
        }
        SharePoint      = @{
            SiteId   = 'theSiteId'
            DriveId  = 'theDriveId'
            FolderId = 'theFolderId'
        }
        Printer         = @{
            Name = 'thePrinterName'
            Port = 9100
        }
        SendMail        = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testData = @(
        @{
            DateTime         = Get-Date
            SiteId           = $testInputFile.SharePoint.SiteId
            DriveId          = $testInputFile.SharePoint.DriveId
            FolderId         = $testInputFile.SharePoint.FolderId
            Printer          = $testInputFile.Printer.Name
            PrinterPort      = $testInputFile.Printer.Port
            FileName         = 'File1.pdf'
            FileCreationDate = Get-Date
            PrintDuplicate   = $true
            Actions          = @('file downloaded', 'file printed')
            Error            = $null
            Printed          = $true
        }
    )

    $testExportedExcelRows = @(
        [PSCustomObject]@{
            DateTime         = $testData[0].DateTime
            FileName         = $testData[0].FileName
            FileCreationDate = $testData[0].FileCreationDate
            PrintDuplicate   = $testData[0].PrintDuplicate
            Printed          = $testData[0].Printed
            Actions          = $testData[0].Actions -join ', '
            Error            = $testData[0].Error
        }
    )

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
    }

    Mock Invoke-PrintFileScriptHC {
        $testData
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
                'SharePoint', 'Printer', 'SendMail'
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
            It 'SharePoint.<_> not found' -ForEach @(
                'SiteId', 'DriveId', 'FolderId'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SharePoint.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SharePoint.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Printer.<_> not found' -ForEach @(
                'Name', 'Port'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Printer.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Printer.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Option.<_> not a boolean' -ForEach @(
                'PrintDuplicate'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Option.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Option.$_' is not a boolean value*")
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
Describe 'when the script runs successfully' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams
    }
    It 'the log folder is created' {
        $testParams.LogFolder | Should -Exist
    }
    It 'the print file script is called' {
        Should -Invoke Invoke-PrintFileScriptHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($SiteId -eq $testInputFile.SharePoint.SiteId) -and
            ($DriveId -eq $testInputFile.SharePoint.DriveId) -and
            ($FolderId -eq $testInputFile.SharePoint.FolderId) -and
            ($PrinterName -eq $testInputFile.Printer.Name) -and
            ($PrinterPort -eq $testInputFile.Printer.Port)
        }
    }
    Context 'create an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter "* - Log.xlsx"

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
                    $_.FileName -eq $testRow.FileName
                }

                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be $testRow.DateTime.ToString('yyyyMMdd')
                $actualRow.FileCreationDate.ToString('yyyyMMdd') |
                Should -Be $testRow.FileCreationDate.ToString('yyyyMMdd')
                $actualRow.PrintDuplicate | Should -Be $testRow.PrintDuplicate
                $actualRow.Printed | Should -Be $testRow.Printed
                $actualRow.Actions | Should -Be $testRow.Actions
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($To -eq $testInputFile.SendMail.To) -and
                ($Priority -eq 'Normal') -and
                ($Subject -eq '1 printed') -and
                ($Attachments -like '*- Log.xlsx') -and
                ($Message -like "*Summary of print actions*table*Printed files*1**")
            }
        }
    }
}
Describe 'Option' {
    Context 'PrintDuplicate' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Option.PrintDuplicate = $true

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams
        }
        Context 'Invoke-PrintFileScriptHC is called' {
            It 'the first time without a last file name printed' {
                Should -Invoke Invoke-PrintFileScriptHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    ($SiteId -eq $testInputFile.SharePoint.SiteId) -and
                    ($DriveId -eq $testInputFile.SharePoint.DriveId) -and
                    ($FolderId -eq $testInputFile.SharePoint.FolderId) -and
                    ($PrinterName -eq $testInputFile.Printer.Name) -and
                    ($PrinterPort -eq $testInputFile.Printer.Port) -and
                    (-not $FileNameLastPrintedFile)
                }
            }
            It 'the second time with a last file name printed' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Option.PrintDuplicate = $false

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Invoke-PrintFileScriptHC -Times 1 -Exactly -ParameterFilter {
                    ($SiteId -eq $testInputFile.SharePoint.SiteId) -and
                    ($DriveId -eq $testInputFile.SharePoint.DriveId) -and
                    ($FolderId -eq $testInputFile.SharePoint.FolderId) -and
                    ($PrinterName -eq $testInputFile.Printer.Name) -and
                    ($PrinterPort -eq $testInputFile.Printer.Port) -and
                    ($FileNameLastPrintedFile -eq $testData[0].FileName)
                }
            }
        }
    }
}
Describe 'Export an Excel file' {
    Context 'create no Excel file' {
        It "when no data is returned from the child script" {
            Mock Invoke-PrintFileScriptHC

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -BeNullOrEmpty
        }
    }
    Context 'create an Excel file' {
        It "'OnlyOnError' and there are errors" {
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $false
            $testNewDate.Error = 'oops'

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
            }

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $true
            $testNewDate.Error = $null

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
            }

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $false
            $testNewDate.Error = 'oops'

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
            }

            $testInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx' |
            Should -Not -BeNullOrEmpty
        }
    }
} -Tag test
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
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $false
            $testNewDate.Error = $null

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
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
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $false
            $testNewDate.Error = 'oops'

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $true
            $testNewDate.Error = $null

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            $testNewDate = Copy-ObjectHC $testData
            $testNewDate.Printed = $false
            $testNewDate.Error = 'oops'

            Mock Invoke-PrintFileScriptHC {
                $testNewDate
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
