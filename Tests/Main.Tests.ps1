BeforeAll {
    $scriptRoot = Split-Path $PSScriptRoot
    . "$scriptRoot\Functions.ps1"
    $main = Join-Path $scriptRoot 'Main.ps1'
}

Describe 'Upload configuration and orchestration' {
    BeforeEach {
        $config = Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json
        $config.Settings.SendMail.When = 'Never'
        $config.Settings.SaveInEventLog.Save = $false
        $config.Settings.SaveLogFiles.Where.Folder = $null
        $config.MaxConcurrent.JobsTotal = 1
        $config.Tasks = @($config.Tasks[0])
        $source = New-Item -ItemType Directory (Join-Path $TestDrive ([guid]::NewGuid()))
        $config.Tasks[0].Folders = @($source.FullName)
        $configPath = Join-Path $TestDrive 'config.json'
        $global:LASTEXITCODE = 0
    }

    It 'accepts the supplied example configuration without running it' {
        $example = Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json
        @(ConvertTo-UploadConfigurationHC $example).Count | Should -Be 2
    }

    It 'exposes only ConfigurationJsonFile as a script-specific parameter' {
        $tokens = $null; $parseErrors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($main, [ref]$tokens, [ref]$parseErrors)
        @($ast.ParamBlock.Parameters.Name.VariablePath.UserPath) | Should -Be @('ConfigurationJsonFile')
    }

    It 'keeps all post-upload actions disabled when omitted' {
        $config.Tasks[0].PSObject.Properties.Remove('AfterUpload')
        $prepared = ConvertTo-UploadConfigurationHC $config
        $prepared.AfterUpload.RemoveFile | Should -BeFalse
        $prepared.AfterUpload.CopyToFolder | Should -BeNullOrEmpty
    }

    It 'rejects string booleans before any uploads' {
        $config.Tasks[0].AfterUpload.RemoveFile = 'false'
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*RemoveFile*'
    }

    It 'rejects copying inside another source folder' {
        $other = Join-Path $TestDrive 'Other'
        $config.Tasks[0].Folders += $other
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $other 'Archive'
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*outside every source*'
    }

    It 'rejects overlapping sources' {
        $config.Tasks[0].Folders += Join-Path $source 'Sub'
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*overlap*'
    }

    It 'rejects unknown task properties' {
        $config.Tasks[0] | Add-Member RemoveAfterUplaod $true
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Unsupported property*'
    }

    It 'runs an empty folder without connecting to SharePoint' {
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        @(& $main $configPath).Count | Should -Be 0
        $global:LASTEXITCODE | Should -Be 0
    }

    It 'uses the bundled uploader and reports a successful upload' {
        Set-Content (Join-Path $source 'report.txt') 'report'
        Import-Module Microsoft.Graph.Authentication
        Mock Get-MgContext { @{ ClientId = 'client'; TenantId = 'tenant' } }
        Mock Connect-MgGraph {}
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'PUT') { return [pscustomobject]@{ id = 'item'; name = 'report.txt'; webUrl = 'https://example.com/report.txt' } }
            if ($Uri -like '*/drives') { return @{ value = @(@{ id = 'drive'; name = 'Documents' }) } }
            return @{ id = 'site' }
        }
        $config.Tasks[0].SharePoint.ClientId = 'client'
        $config.Tasks[0].SharePoint.TenantId = 'tenant'
        $config.Tasks[0].SharePoint.CertificateThumbprint = 'cert'
        $config.Tasks[0].SharePoint.FolderPath = ''
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Uploaded | Should -BeTrue
        $result.Error | Should -BeNullOrEmpty
        $global:LASTEXITCODE | Should -Be 0
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter { $Method -eq 'PUT' }
    }

    It 'fails before execution when a later task is invalid' {
        $config.Tasks += [pscustomobject]@{ Folders = @('C:\Invalid'); Recurse = 'false' }
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        & $main $configPath -WarningVariable warnings
        $global:LASTEXITCODE | Should -Be 1
        "$warnings" | Should -BeLike '*Recurse*'
    }

    It 'returns a failed result and exit 1 for a missing source' {
        $config.Tasks[0].Folders = @(Join-Path $TestDrive 'missing')
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Not -BeNullOrEmpty
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'writes a JSON run log for an empty source' {
        $config.Settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive 'Logs'
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        & $main $configPath
        $log = Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '* - Log.json'
        $run = Get-Content -LiteralPath $log.FullName -Raw | ConvertFrom-Json
        $run.Uploaded | Should -Be 0
        $run.Errors | Should -Be 0
        $global:LASTEXITCODE | Should -Be 0
    }

    It 'runs multiple folder jobs in parallel without dropping results' {
        $config.MaxConcurrent.JobsTotal = 3
        $config.MaxConcurrent.JobsPerComputer = 2
        $config.Tasks[0].Folders = @(1..5 | ForEach-Object { Join-Path $TestDrive "missing-$_" })
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $results = @(& $main $configPath)
        $results.Count | Should -Be 5
        @($results | Where-Object Error).Count | Should -Be 5
        @($results.Source | Sort-Object -Unique).Count | Should -Be 5
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'sends the worker and uploader to the configured remote endpoint and closes the session' {
        $config.Tasks[0].ComputerName = 'SERVER1'
        $config.PSSessionConfiguration = 'Custom.PowerShell7'
        Mock New-PSSession { New-MockObject -Type 'System.Management.Automation.Runspaces.PSSession' }
        Mock Invoke-Command { throw 'Remote upload failed' }
        Mock Remove-PSSession {}
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Be 'Remote upload failed'
        Should -Invoke New-PSSession -Times 1 -ParameterFilter { $ComputerName -eq 'SERVER1' -and $ConfigurationName -eq 'Custom.PowerShell7' }
        Should -Invoke Invoke-Command -Times 1 -ParameterFilter { $FilePath -like '*Upload files.ps1' -and $ArgumentList[1] -like '*function Send-GraphFileHC*' }
        Should -Invoke Remove-PSSession -Times 1
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'reports remote-session creation failure without retrying the whole job' {
        $config.Tasks[0].ComputerName = 'SERVER1'
        Mock New-PSSession { throw 'Access denied' }
        Mock Invoke-Command {}
        Mock Remove-PSSession {}
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Be 'Access denied'
        Should -Invoke New-PSSession -Times 1
        Should -Invoke Invoke-Command -Times 0
        Should -Invoke Remove-PSSession -Times 0
    }
}

Describe 'Upload reporting' {
    BeforeEach {
        $settings = (Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json).Settings
        $settings.SaveInEventLog.Save = $false
        $settings.SaveLogFiles.Where.Folder = $null
        $systemErrors = [Collections.Generic.List[object]]::new()
        $row = [pscustomobject]@{
            ComputerName = 'localhost'; Name = 'Reports'; Source = 'C:\Exports\report.txt'; Destination = 'Reports/report.txt'
            Uploaded = $true; CopiedTo = ''; Removed = $false; Error = ''
        }
        Mock Send-MailKitMessageHC {}
        Mock Write-EventsToEventLogHC {}
    }

    It 'applies mail condition <When> with Uploaded=<Uploaded> and Error=<HasError>' -ForEach @(
        @{ When = 'Never'; Uploaded = $true; HasError = $true; Expected = 0 }
        @{ When = 'Always'; Uploaded = $false; HasError = $false; Expected = 1 }
        @{ When = 'OnError'; Uploaded = $true; HasError = $false; Expected = 0 }
        @{ When = 'OnError'; Uploaded = $false; HasError = $true; Expected = 1 }
        @{ When = 'OnErrorOrAction'; Uploaded = $false; HasError = $false; Expected = 0 }
        @{ When = 'OnErrorOrAction'; Uploaded = $true; HasError = $false; Expected = 1 }
        @{ When = 'OnErrorOrAction'; Uploaded = $false; HasError = $true; Expected = 1 }
    ) {
        $settings.SendMail.When = $When
        $row.Uploaded = $Uploaded
        if ($HasError) { $row.Error = 'Failed' }
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        Should -Invoke Send-MailKitMessageHC -Times $Expected -Exactly
        $systemErrors.Count | Should -Be 0
    }

    It 'counts a post-upload failure as both an upload and an error' {
        $row.Error = 'Archive already exists'
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        Should -Invoke Send-MailKitMessageHC -Times 1 -ParameterFilter { $Subject -eq '1 uploaded, 1 errors' -and $Body -like '*Archive already exists*' }
    }

    It 'records SMTP failures as system errors' {
        Mock Send-MailKitMessageHC { throw 'SMTP unavailable' }
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 1
        $systemErrors[0].Message | Should -BeLike '*SMTP unavailable*'
    }

    It 'writes Excel results, JSON details and the email HTML' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        Import-Module ImportExcel
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 0
        $files = @(Get-ChildItem $settings.SaveLogFiles.Where.Folder)
        $files.Count | Should -Be 3
        $excel = $files | Where-Object Extension -eq '.xlsx'
        (Import-Excel $excel.FullName).Source | Should -Be $row.Source
        $html = $files | Where-Object Extension -eq '.html'
        $script:expectedHtml = Get-Content $html.FullName -Raw
        Should -Invoke Send-MailKitMessageHC -Times 1 -ParameterFilter { $Body.TrimEnd() -eq $script:expectedHtml.TrimEnd() -and $Attachments.Count -eq 1 }
    }

    It 'removes only old generated logs belonging to this script and configuration' {
        $folder = New-Item -ItemType Directory (Join-Path $TestDrive ([guid]::NewGuid()))
        $settings.SaveLogFiles.Where.Folder = $folder.FullName
        $settings.SaveLogFiles.DeleteLogsAfterDays = 30
        $settings.SendMail.When = 'Never'
        $oldName = '2020_01_01_120000_000 - Upload files to SharePoint (Test) - Log.json'
        $otherName = '2020_01_01_120000_000 - Upload files to SharePoint (Other) - Log.json'
        $unrelatedName = '2020_01_01_120000_000 - Upload files to SharePoint (Test) - keep.txt'
        foreach ($name in $oldName, $otherName, $unrelatedName) {
            $file = New-Item -ItemType File (Join-Path $folder $name)
            $file.LastWriteTime = (Get-Date).AddDays(-60)
        }
        Send-UploadReportHC $settings @() $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path $folder $oldName) | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $folder $otherName) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $folder $unrelatedName) | Should -BeTrue
    }

    It 'records event-log failures in the JSON run log' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        $settings.SaveInEventLog.Save = $true
        Mock Write-EventsToEventLogHC { throw 'Event log unavailable' }
        Send-UploadReportHC $settings @() $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 1
        $json = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.json'
        $run = Get-Content $json.FullName -Raw | ConvertFrom-Json
        $run.Errors | Should -Be 1
        $run.SystemErrors[0].Message | Should -BeLike '*Event log unavailable*'
    }
}