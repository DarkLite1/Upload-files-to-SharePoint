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
        $config.Tasks[0].Upload.Folders = @($source.FullName)
        $configPath = Join-Path $TestDrive 'config.json'
        $global:LASTEXITCODE = 0
    }

    It 'accepts the supplied example configuration without running it' {
        $example = Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json
        $prepared = @(ConvertTo-UploadConfigurationHC $example)
        $prepared.Count | Should -Be 2
        $prepared.UploadType | Should -Be @('Folder', 'File')
    }

    It 'rejects the old top-level Folders property' {
        $config.Tasks[0] | Add-Member -NotePropertyName Folders -NotePropertyValue @($source.FullName)
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Unsupported property*Folders*'
    }

    It 'requires Upload instead of silently accepting a task without sources' {
        $config.Tasks[0].PSObject.Properties.Remove('Upload')
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Upload must be an object*'
    }

    It 'requires Recurse when folders are selected' {
        $config.Tasks[0].PSObject.Properties.Remove('Recurse')
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Recurse*'
    }

    It 'rejects non-Boolean Recurse even for a file-only task' {
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @(Join-Path $source.FullName 'report.txt') }
        $config.Tasks[0].Recurse = 'false'
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Recurse*'
    }

    It 'accepts file-only tasks without Recurse and preserves the optional display name' {
        $path = Join-Path $source.FullName 'report[1].txt'
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @([pscustomobject]@{ Name = 'Selected report'; Path = $path }) }
        $config.Tasks[0].PSObject.Properties.Remove('Recurse')
        $prepared = ConvertTo-UploadConfigurationHC $config
        $prepared.UploadType | Should -Be 'File'
        $prepared.Path | Should -Be $path
        $prepared.Name | Should -Be 'Selected report'
        $prepared.Recurse | Should -BeFalse
    }

    It 'prepares mixed selections as separate folder and file jobs' {
        $path = Join-Path $TestDrive 'selected.txt'
        $config.Tasks[0].Upload.Files = @($path)
        $prepared = @(ConvertTo-UploadConfigurationHC $config)
        $prepared.Count | Should -Be 2
        $prepared.UploadType | Should -Be @('Folder', 'File')
        $prepared[0].Recurse | Should -BeTrue
        $prepared[1].Recurse | Should -BeFalse
        $prepared[1].Path | Should -Be $path
    }

    It 'rejects invalid Upload selections: <Description>' -ForEach @(
        @{ Description = 'null'; Upload = $null }
        @{ Description = 'scalar'; Upload = 'C:\Reports' }
        @{ Description = 'empty object'; Upload = [pscustomobject]@{} }
        @{ Description = 'empty arrays'; Upload = [pscustomobject]@{ Folders = @(); Files = @() } }
        @{ Description = 'unknown property'; Upload = [pscustomobject]@{ File = @('C:\report.txt') } }
        @{ Description = 'null files'; Upload = [pscustomobject]@{ Files = $null } }
        @{ Description = 'null folders'; Upload = [pscustomobject]@{ Folders = $null } }
        @{ Description = 'scalar folders'; Upload = [pscustomobject]@{ Folders = 'C:\Reports' } }
        @{ Description = 'scalar files'; Upload = [pscustomobject]@{ Files = 'C:\report.txt' } }
        @{ Description = 'blank file'; Upload = [pscustomobject]@{ Files = @(' ') } }
        @{ Description = 'relative file'; Upload = [pscustomobject]@{ Files = @('report.txt') } }
        @{ Description = 'wildcard file'; Upload = [pscustomobject]@{ Files = @('C:\*.txt') } }
        @{ Description = 'wildcard folder'; Upload = [pscustomobject]@{ Folders = @('C:\Reports*') } }
        @{ Description = 'invalid entry'; Upload = [pscustomobject]@{ Files = @(42) } }
        @{ Description = 'missing Path'; Upload = [pscustomobject]@{ Files = @([pscustomobject]@{ Name = 'Report' }) } }
        @{ Description = 'invalid Name'; Upload = [pscustomobject]@{ Files = @([pscustomobject]@{ Name = 42; Path = 'C:\report.txt' }) } }
        @{ Description = 'unknown entry property'; Upload = [pscustomobject]@{ Files = @([pscustomobject]@{ Path = 'C:\report.txt'; Rename = 'other.txt' }) } }
    ) {
        $config.Tasks[0].Upload = $Upload
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Upload*'
    }

    It 'rejects selecting a file also covered by an uploaded folder' {
        $config.Tasks[0].Upload.Files = @(Join-Path $source.FullName 'report.txt')
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*overlap*'
    }

    It 'rejects duplicate explicit file paths after normalization' {
        $path = Join-Path $source.FullName 'report.txt'
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @($path, (Join-Path $source.FullName 'Sub/../REPORT.txt')) }
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*overlap*'
    }

    It 'rejects an archive containing an explicitly selected source file' {
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @(Join-Path $source.FullName 'report.txt') }
        $config.Tasks[0].AfterUpload.CopyToFolder = $source.FullName
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*must not contain a selected source file*'
    }

    It 'allows an archive beside a selected file without selecting its parent folder' {
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @(Join-Path $source.FullName 'report.txt') }
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $source.FullName 'Archive'
        (ConvertTo-UploadConfigurationHC $config).AfterUpload.CopyToFolder | Should -Be $config.Tasks[0].AfterUpload.CopyToFolder
    }

    It 'routes an explicit UNC file on the script host' {
        $config.Tasks[0].ComputerName = $null
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @('//server/share/Sub/../report[1].txt') }
        $prepared = ConvertTo-UploadConfigurationHC $config
        $prepared.UploadType | Should -Be 'File'
        $prepared.Path | Should -Be '\\server\share\report[1].txt'
        $prepared.ComputerName | Should -Be $env:COMPUTERNAME
        $prepared.Target | Should -Be 'server'
    }

    It 'rejects a remote UNC file to avoid a double hop' {
        $config.Tasks[0].ComputerName = 'SERVER1'
        $config.Tasks[0].Upload = [pscustomobject]@{ Files = @('\\server\share\report.txt') }
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*double hop*'
    }

    It 'resolves ENV paths for selected files on the script host' {
        $variable = 'UPLOAD_SELECTION_TEST_PATH'
        $previous = [Environment]::GetEnvironmentVariable($variable)
        $path = Join-Path $source.FullName 'report.txt'
        try {
            [Environment]::SetEnvironmentVariable($variable, $path)
            $config.Tasks[0].Upload = [pscustomobject]@{ Files = @("ENV:$variable") }
            (ConvertTo-UploadConfigurationHC $config).Path | Should -Be $path
        }
        finally { [Environment]::SetEnvironmentVariable($variable, $previous) }
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

    It 'accepts Boolean overwrite settings' -ForEach @($false, $true) {
        $config.Tasks[0].SharePoint | Add-Member -NotePropertyName Overwrite -NotePropertyValue $_ -Force
        (ConvertTo-UploadConfigurationHC $config).SharePoint.Overwrite | Should -Be $_
    }

    It 'rejects non-Boolean overwrite settings' -ForEach @('false', 1, $null) {
        $config.Tasks[0].SharePoint | Add-Member -NotePropertyName Overwrite -NotePropertyValue $_ -Force
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Overwrite must be true or false*'
    }

    It 'accepts <Description> attribute exclusions' -ForEach @(
        @{ Description = 'omitted'; Exclude = $null; Omit = $true; Expected = @() }
        @{ Description = 'empty object'; Exclude = [pscustomobject]@{}; Omit = $false; Expected = @() }
        @{ Description = 'empty array'; Exclude = [pscustomobject]@{ Attributes = @() }; Omit = $false; Expected = @() }
        @{ Description = 'Hidden'; Exclude = [pscustomobject]@{ Attributes = @('Hidden') }; Omit = $false; Expected = @('Hidden') }
        @{ Description = 'System'; Exclude = [pscustomobject]@{ Attributes = @('System') }; Omit = $false; Expected = @('System') }
        @{ Description = 'both attributes'; Exclude = [pscustomobject]@{ Attributes = @('Hidden', 'System') }; Omit = $false; Expected = @('Hidden', 'System') }
        @{ Description = 'duplicate attributes'; Exclude = [pscustomobject]@{ Attributes = @('Hidden', 'Hidden') }; Omit = $false; Expected = @('Hidden') }
    ) {
        if ($Omit) { $config.Tasks[0].PSObject.Properties.Remove('Exclude') }
        else { $config.Tasks[0].Exclude = $Exclude }
        $prepared = ConvertTo-UploadConfigurationHC $config
        @($prepared.Exclude.Attributes).Count | Should -Be $Expected.Count
        ($prepared.Exclude.Attributes -join ',') | Should -Be ($Expected -join ',')
    }

    It 'rejects invalid exclusions: <Description>' -ForEach @(
        @{ Description = 'null object'; Exclude = $null }
        @{ Description = 'scalar object'; Exclude = 'Hidden' }
        @{ Description = 'null array'; Exclude = [pscustomobject]@{ Attributes = $null } }
        @{ Description = 'scalar attribute'; Exclude = [pscustomobject]@{ Attributes = 'Hidden' } }
        @{ Description = 'unsupported attribute'; Exclude = [pscustomobject]@{ Attributes = @('ReadOnly') } }
        @{ Description = 'numeric attribute'; Exclude = [pscustomobject]@{ Attributes = @(2) } }
        @{ Description = 'null attribute'; Exclude = [pscustomobject]@{ Attributes = @($null) } }
        @{ Description = 'unknown property'; Exclude = [pscustomobject]@{ Attribute = @('Hidden') } }
    ) {
        $config.Tasks[0].Exclude = $Exclude
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*Exclude*'
    }

    It 'passes exclusions to the <Selection> worker with JobsTotal <JobsTotal>' -ForEach @(
        @{ JobsTotal = 1; Selection = 'Folders' }
        @{ JobsTotal = 2; Selection = 'Folders' }
        @{ JobsTotal = 1; Selection = 'Files' }
        @{ JobsTotal = 2; Selection = 'Files' }
    ) {
        $path = Join-Path $source 'hidden.txt'
        Set-Content -LiteralPath $path -Value 'protected'
        [IO.File]::SetAttributes($path, [IO.FileAttributes]::Hidden)
        if ($Selection -eq 'Files') { $config.Tasks[0].Upload = [pscustomobject]@{ Files = @($path) } }
        $config.MaxConcurrent.JobsTotal = $JobsTotal
        $config.Tasks[0].Exclude.Attributes = @('Hidden')
        $config.Tasks[0].AfterUpload.RemoveFile = $true
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $TestDrive ([guid]::NewGuid())
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        @(& $main $configPath).Count | Should -Be 0
        $global:LASTEXITCODE | Should -Be 0
        Test-Path -LiteralPath $path | Should -BeTrue
        Test-Path -LiteralPath $config.Tasks[0].AfterUpload.CopyToFolder | Should -BeFalse
    }

    It 'accepts empty file and folder exclusions' {
        $config.Tasks[0].Exclude = [pscustomobject]@{ Attributes = @(); Folders = @(); Files = @() }
        $prepared = ConvertTo-UploadConfigurationHC $config
        @($prepared.Exclude.Folders).Count | Should -Be 0
        @($prepared.Exclude.Files).Count | Should -Be 0
    }

    It 'routes each exclusion to its source folder after normalization' {
        $other = Join-Path $TestDrive 'OtherReports'
        $config.Tasks[0].Upload.Folders += $other
        $config.Tasks[0].Exclude = [pscustomobject]@{
            Folders = @((Join-Path $source 'Unused/../Drafts/').Replace('\', '/'))
            Files = @((Join-Path $other 'status.json').ToUpperInvariant())
        }
        $prepared = @(ConvertTo-UploadConfigurationHC $config)
        $prepared.Count | Should -Be 2
        $prepared[0].Exclude.Folders | Should -Be @(Join-Path $source 'Drafts')
        @($prepared[0].Exclude.Files).Count | Should -Be 0
        @($prepared[1].Exclude.Folders).Count | Should -Be 0
        $prepared[1].Exclude.Files | Should -Be @((Join-Path $other 'status.json').ToUpperInvariant())
    }

    It 'accepts UNC exclusions below the correct source share' {
        $config.Tasks[0].ComputerName = $null
        $config.Tasks[0].Upload.Folders = @('\\server\share\Reports')
        $config.Tasks[0].Exclude = [pscustomobject]@{
            Folders = @('\\SERVER\share\Reports\Drafts\')
            Files = @('\\server\share\Reports\status.json')
        }
        $prepared = ConvertTo-UploadConfigurationHC $config
        $prepared.Exclude.Folders | Should -Be @('\\SERVER\share\Reports\Drafts')
        $prepared.Exclude.Files | Should -Be @('\\server\share\Reports\status.json')
    }

    It 'rejects <Case> in Exclude.<Property>' -ForEach @(
        foreach ($property in 'Folders', 'Files') {
            foreach ($case in 'Null', 'Scalar', 'Blank', 'Number', 'Relative', 'Root', 'Sibling', 'Traversal', 'Wildcard') {
                @{ Property = $property; Case = $case }
            }
        }
    ) {
        $values = switch ($Case) {
            'Null' { $null }
            'Scalar' { Join-Path $source 'entry' }
            'Blank' { ,@(' ') }
            'Number' { ,@(1) }
            'Relative' { ,@('entry') }
            'Root' { ,@($source.FullName) }
            'Sibling' { ,@($source.FullName + 'Other\entry') }
            'Traversal' { ,@(Join-Path $source '..\outside') }
            'Wildcard' { ,@(Join-Path $source '*.txt') }
        }
        $config.Tasks[0].Exclude = [pscustomobject]@{ $Property = $values }
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw "*Exclude.$Property*"
    }

    It 'passes a file exclusion for <Selection> through Main without post-upload actions' -ForEach @('Folders', 'Files') {
        $selection = $_
        $path = Join-Path $source 'protected.txt'
        Set-Content -LiteralPath $path -Value 'protected'
        if ($selection -eq 'Files') { $config.Tasks[0].Upload = [pscustomobject]@{ Files = @($path) } }
        $config.Tasks[0].Exclude = [pscustomobject]@{ Files = @($path.ToUpperInvariant().Replace('\', '/')) }
        $config.Tasks[0].AfterUpload.RemoveFile = $true
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $TestDrive ([guid]::NewGuid())
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        @(& $main $configPath).Count | Should -Be 0
        $global:LASTEXITCODE | Should -Be 0
        Test-Path -LiteralPath $path | Should -BeTrue
        Test-Path -LiteralPath $config.Tasks[0].AfterUpload.CopyToFolder | Should -BeFalse
    }

    It 'rejects copying inside another source folder' {
        $other = Join-Path $TestDrive 'Other'
        $config.Tasks[0].Upload.Folders += $other
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $other 'Archive'
        { ConvertTo-UploadConfigurationHC $config } | Should -Throw '*outside every source*'
    }

    It 'rejects overlapping sources' {
        $config.Tasks[0].Upload.Folders += Join-Path $source 'Sub'
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

    It 'passes overwrite <Overwrite> for Upload.<Selection> to the bundled uploader' -ForEach @(
        @{ Overwrite = $false; Behavior = 'fail'; Selection = 'Folders' }
        @{ Overwrite = $true; Behavior = 'replace'; Selection = 'Folders' }
        @{ Overwrite = $false; Behavior = 'fail'; Selection = 'Files' }
        @{ Overwrite = $true; Behavior = 'replace'; Selection = 'Files' }
    ) {
        $path = Join-Path $source 'report.txt'
        Set-Content -LiteralPath $path -Value 'report'
        if ($Selection -eq 'Files') {
            $config.Tasks[0].Upload = [pscustomobject]@{ Files = @($path) }
            $config.Tasks[0].PSObject.Properties.Remove('Recurse')
        }
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
        $config.Tasks[0].SharePoint.Overwrite = $Overwrite
        $config.Tasks[0].AfterUpload.RemoveFile = $true
        $config.Tasks[0].AfterUpload.CopyToFolder = Join-Path $TestDrive ([guid]::NewGuid())
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Uploaded | Should -BeTrue
        $result.Error | Should -BeNullOrEmpty
        $result.Action | Should -Be 'Uploaded, Copied, Removed'
        $result.Destination | Should -Be 'report.txt'
        $result.FullName | Should -Be $path
        Test-Path -LiteralPath $path | Should -BeFalse
        Get-Content -LiteralPath $result.CopiedTo | Should -Be 'report'
        $global:LASTEXITCODE | Should -Be 0
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and "$Uri".EndsWith("/content?@microsoft.graph.conflictBehavior=$Behavior")
        }
    }

    It 'fails before execution when a later task is invalid' {
        $config.Tasks += [pscustomobject]@{ Upload = [pscustomobject]@{ Folders = @('C:\Invalid') }; Recurse = 'false' }
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        & $main $configPath -WarningVariable warnings
        $global:LASTEXITCODE | Should -Be 1
        "$warnings" | Should -BeLike '*Recurse*'
    }

    It 'returns a failed result and exit 1 for a missing source' {
        $config.Tasks[0].Upload.Folders = @(Join-Path $TestDrive 'missing')
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Not -BeNullOrEmpty
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'does not create Excel or system-errors JSON when the source is <Description>' -ForEach @(
        @{ Description = 'empty'; ExcludeFile = $false; DirectFile = $false }
        @{ Description = 'fully excluded'; ExcludeFile = $true; DirectFile = $false }
        @{ Description = 'an excluded explicit file'; ExcludeFile = $true; DirectFile = $true }
    ) {
        if ($ExcludeFile) {
            $excludedPath = Join-Path $source.FullName 'excluded.txt'
            Set-Content -LiteralPath $excludedPath -Value 'excluded report'
            $config.Tasks[0].Exclude.Files = @($excludedPath)
            if ($DirectFile) { $config.Tasks[0].Upload = [pscustomobject]@{ Files = @($excludedPath) } }
        }
        $config.Settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        @(& $main $configPath).Count | Should -Be 0
        @(Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '*.xlsx').Count | Should -Be 0
        @(Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '*.json').Count | Should -Be 0
        $global:LASTEXITCODE | Should -Be 0
    }

    It 'logs configuration errors only in the system-errors JSON file' {
        $config.Settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive 'ConfigErrors'
        $config.Tasks[0].SharePoint.Overwrite = 'false'
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        & $main $configPath
        $json = Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '* - System errors log.json'
        $errors = @(Get-Content -LiteralPath $json.FullName -Raw | ConvertFrom-Json)
        $errors.Count | Should -Be 1
        $errors[0].Message | Should -BeLike '*Overwrite must be true or false*'
        @(Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '*.xlsx').Count | Should -Be 0
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'runs multiple folder jobs in parallel without dropping results' {
        $config.MaxConcurrent.JobsTotal = 3
        $config.MaxConcurrent.JobsPerComputer = 2
        $config.Tasks[0].Upload.Folders = @(1..5 | ForEach-Object { Join-Path $TestDrive "missing-$_" })
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $results = @(& $main $configPath)
        $results.Count | Should -Be 5
        @($results | Where-Object Error).Count | Should -Be 5
        @($results.Source | Sort-Object -Unique).Count | Should -Be 5
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'sends a <UploadType> job to the remote endpoint and closes the session' -ForEach @('Folder', 'File') {
        $uploadType = $_
        if ($uploadType -eq 'File') {
            $config.Tasks[0].Upload = [pscustomobject]@{ Files = @(Join-Path $source.FullName 'report.txt') }
        }
        $config.Tasks[0].ComputerName = 'SERVER1'
        $config.PSSessionConfiguration = 'Custom.PowerShell7'
        Mock New-PSSession { New-MockObject -Type 'System.Management.Automation.Runspaces.PSSession' }
        Mock Invoke-Command { throw 'Remote upload failed' }
        Mock Remove-PSSession {}
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Be 'Remote upload failed'
        Should -Invoke New-PSSession -Times 1 -ParameterFilter { $ComputerName -eq 'SERVER1' -and $ConfigurationName -eq 'Custom.PowerShell7' }
        Should -Invoke Invoke-Command -Times 1 -ParameterFilter {
            $FilePath -like '*Upload files.ps1' -and $ArgumentList[0].UploadType -eq $uploadType -and $ArgumentList[1] -like '*function Send-GraphFileHC*'
        }
        Should -Invoke Remove-PSSession -Times 1
        $global:LASTEXITCODE | Should -Be 1
    }

    It 'reports remote-session creation failure without retrying the whole job' {
        $config.Tasks[0].ComputerName = 'SERVER1'
        $config.Settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        Mock New-PSSession { throw 'Access denied' }
        Mock Invoke-Command {}
        Mock Remove-PSSession {}
        $config | ConvertTo-Json -Depth 12 | Set-Content $configPath
        $result = & $main $configPath
        $result.Error | Should -Be 'Access denied'
        $result.Stage | Should -Be 'Open remote session'
        $excel = Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '*.xlsx'
        $errorRow = Import-Excel -Path $excel.FullName -WorksheetName 'Errors'
        $errorRow.FullName | Should -Be $source.FullName
        $errorRow.Stage | Should -Be 'Open remote session'
        $errorRow.Error | Should -Be 'Access denied'
        $errorRow.ExceptionType | Should -Not -BeNullOrEmpty
        $errorRow.FullyQualifiedErrorId | Should -Not -BeNullOrEmpty
        $errorRow.ScriptStackTrace | Should -Not -BeNullOrEmpty
        @(Get-ChildItem $config.Settings.SaveLogFiles.Where.Folder -Filter '*.json').Count | Should -Be 0
        Should -Invoke New-PSSession -Times 1
        Should -Invoke Invoke-Command -Times 0
        Should -Invoke Remove-PSSession -Times 0
    }
}

Describe 'SharePoint overwrite behavior' {
    BeforeEach {
        Import-Module Microsoft.Graph.Authentication
        $uploadPath = Join-Path $TestDrive 'report.txt'
        Set-Content -LiteralPath $uploadPath -Value 'report'
        $uploadParams = @{
            FilePath = $uploadPath; SiteUrl = 'https://contoso.sharepoint.com/sites/IT'
            DocumentLibraryName = 'Documents'; ClientId = 'client'; TenantId = 'tenant'
            CertificateThumbprint = 'cert'; MaxRetries = 1
        }
        Mock Get-MgContext { @{ ClientId = 'client'; TenantId = 'tenant' } }
        Mock Connect-MgGraph {}
        Mock Invoke-MgGraphRequest {
            if ($Method -eq 'PUT') { return @{ id = 'item'; name = 'report.txt' } }
            if ($Method -eq 'POST') { return @{ uploadUrl = 'https://upload.example.com/session' } }
            if ($Uri -like '*/drives') { return @{ value = @(@{ id = 'drive'; name = 'Documents' }) } }
            return @{ id = 'site' }
        }
        Mock Invoke-RestMethod { @{ id = 'item'; name = 'report.txt' } }
        Mock Start-Sleep {}
    }

    It 'defaults to server-side fail for a small file' {
        $result = & "$scriptRoot\UploadToSharePoint.ps1" @uploadParams
        $result.Id | Should -Be 'item'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and "$Uri".EndsWith('/content?@microsoft.graph.conflictBehavior=fail')
        }
    }

    It 'uses <Behavior> for small uploads with overwrite <Overwrite>' -ForEach @(
        @{ Overwrite = $false; Behavior = 'fail' }
        @{ Overwrite = $true; Behavior = 'replace' }
    ) {
        $uploadParams.Overwrite = $Overwrite
        & "$scriptRoot\UploadToSharePoint.ps1" @uploadParams
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and "$Uri".EndsWith("/content?@microsoft.graph.conflictBehavior=$Behavior")
        }
    }

    It 'also protects empty files against overwriting' {
        Clear-Content -LiteralPath $uploadPath
        & "$scriptRoot\UploadToSharePoint.ps1" @uploadParams
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'PUT' -and "$Uri".EndsWith('/content?@microsoft.graph.conflictBehavior=fail')
        }
    }

    It 'sets <Behavior> in large-file upload sessions' -ForEach @(
        @{ Overwrite = $false; Behavior = 'fail' }
        @{ Overwrite = $true; Behavior = 'replace' }
    ) {
        $stream = [IO.File]::OpenWrite($uploadPath)
        try { $stream.SetLength(6MB) } finally { $stream.Dispose() }
        $uploadParams.Overwrite = $Overwrite
        $result = & "$scriptRoot\UploadToSharePoint.ps1" @uploadParams
        $result.Id | Should -Be 'item'
        Should -Invoke Invoke-MgGraphRequest -Times 1 -ParameterFilter {
            $Method -eq 'POST' -and ($Body | ConvertFrom-Json).item.'@microsoft.graph.conflictBehavior' -eq $Behavior
        }
        Should -Invoke Invoke-RestMethod -Times 2 -ParameterFilter { $Method -eq 'PUT' }
    }

    It 'reports an existing destination without retrying or post-upload actions' {
        Mock Invoke-MgGraphRequest { throw '409 nameAlreadyExists' } -ParameterFilter { $Method -eq 'PUT' }
        $config = Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json
        $task = [pscustomobject]@{
            TaskIndex = 0; ComputerName = 'localhost'; Name = 'Reports'; Path = $TestDrive; Recurse = $false
            SharePoint = $config.Tasks[0].SharePoint
            AfterUpload = [pscustomobject]@{ RemoveFile = $true; CopyToFolder = $null }
        }
        $task.SharePoint.ClientId = 'client'
        $task.SharePoint.TenantId = 'tenant'
        $task.SharePoint.CertificateThumbprint = 'cert'
        $task.SharePoint.FolderPath = ''
        $uploader = Get-Content "$scriptRoot\UploadToSharePoint.ps1" -Raw
        $result = & "$scriptRoot\Upload files.ps1" $task $uploader
        $result.Error | Should -BeLike '*already exists and overwriting is disabled*'
        $result.Uploaded | Should -BeFalse
        $result.Removed | Should -BeFalse
        Test-Path -LiteralPath $uploadPath | Should -BeTrue
        Should -Invoke Start-Sleep -Times 0
    }

    It 'cleans up a large-file session when a conflict occurs on the final chunk' {
        $stream = [IO.File]::OpenWrite($uploadPath)
        try { $stream.SetLength(6MB) } finally { $stream.Dispose() }
        Mock Invoke-RestMethod { throw '409 nameAlreadyExists' } -ParameterFilter {
            $Method -eq 'PUT' -and $Headers.'Content-Range' -like 'bytes 5242880-*'
        }
        { & "$scriptRoot\UploadToSharePoint.ps1" @uploadParams } | Should -Throw '*already exists and overwriting is disabled*'
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter { $Method -eq 'DELETE' }
        Should -Invoke Start-Sleep -Times 0
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

    It 'writes Excel results and email HTML without an Errors sheet or JSON file on success' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        Import-Module ImportExcel
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 0
        $files = @(Get-ChildItem $settings.SaveLogFiles.Where.Folder)
        $files.Count | Should -Be 2
        @($files | Where-Object Extension -eq '.json').Count | Should -Be 0
        $excel = $files | Where-Object Extension -eq '.xlsx'
        (Import-Excel $excel.FullName).Source | Should -Be $row.Source
        $package = Open-ExcelPackage -Path $excel.FullName
        try {
            $package.Workbook.Worksheets.Name | Should -Contain 'Overview'
            $package.Workbook.Worksheets.Name | Should -Not -Contain 'Errors'
        }
        finally { $package.Dispose() }
        $html = $files | Where-Object Extension -eq '.html'
        $script:expectedHtml = Get-Content $html.FullName -Raw
        Should -Invoke Send-MailKitMessageHC -Times 1 -ParameterFilter { $Body.TrimEnd() -eq $script:expectedHtml.TrimEnd() -and $Attachments.Count -eq 1 }
    }

    It 'removes only old generated logs belonging to this script and configuration' {
        $folder = New-Item -ItemType Directory (Join-Path $TestDrive ([guid]::NewGuid()))
        $settings.SaveLogFiles.Where.Folder = $folder.FullName
        $settings.SaveLogFiles.DeleteLogsAfterDays = 30
        $settings.SendMail.When = 'Never'
        $oldName = '2020_01_01_120000_000 - Upload files to SharePoint (Test) - System errors log.json'
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

    It 'records event-log failures in the system-errors JSON file' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        $settings.SaveInEventLog.Save = $true
        Mock Write-EventsToEventLogHC { throw 'Event log unavailable' }
        Send-UploadReportHC $settings @() $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 1
        $json = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.json'
        $errors = @(Get-Content $json.FullName -Raw | ConvertFrom-Json)
        $errors.Count | Should -Be 1
        $errors[0].Message | Should -BeLike '*Event log unavailable*'
    }

    It 'logs every file result in Excel and only system errors in JSON' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        $conflict = [pscustomobject]@{
            ComputerName = 'localhost'; Name = 'Reports'; Source = 'C:\Exports\existing.txt'; Destination = 'Reports/existing.txt'
            Uploaded = $false; CopiedTo = ''; Removed = $false; Error = 'SharePoint file already exists and overwriting is disabled.'
        }
        $copyError = [pscustomobject]@{
            ComputerName = 'localhost'; Name = 'Reports'; Source = 'C:\Exports\archive.txt'; Destination = 'Reports/archive.txt'
            Uploaded = $true; CopiedTo = ''; Removed = $false; Error = 'Archive already exists'
        }
        $systemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = 'Configuration failure' })
        Send-UploadReportHC $settings @($row, $conflict, $copyError) $systemErrors (Get-Date) 'Test'
        $excel = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.xlsx'
        $logged = @(Import-Excel $excel.FullName)
        $logged.Count | Should -Be 3
        @($logged | Where-Object Uploaded).Count | Should -Be 2
        @($logged | Where-Object Error).Count | Should -Be 2
        ($logged | Where-Object Source -eq $conflict.Source).Error | Should -Be $conflict.Error
        ($logged | Where-Object Source -eq $copyError.Source).Error | Should -Be $copyError.Error
        $errorRows = @(Import-Excel -Path $excel.FullName -WorksheetName 'Errors')
        $errorRows.Count | Should -Be 2
        $errorRows.Error | Should -Contain $conflict.Error
        $errorRows.Error | Should -Contain $copyError.Error
        $errorRows.Error | Should -Not -Contain 'Configuration failure'
        $json = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '* - System errors log.json'
        $text = Get-Content $json.FullName -Raw
        $errors = @($text | ConvertFrom-Json)
        $errors.Count | Should -Be 1
        $errors[0].Message | Should -Be 'Configuration failure'
        $errors[0].DateTime | Should -Not -BeNullOrEmpty
        $text | Should -Not -Match 'Source|Uploaded|Destination|Archive already exists'
    }

    It 'exports real file dates, paths, size and actions after the source has been removed' {
        $sourceFolder = New-Item -ItemType Directory (Join-Path $TestDrive ([guid]::NewGuid()))
        $path = Join-Path $sourceFolder '000123.txt'
        Set-Content -LiteralPath $path -Value 'report'
        $sourceFile = Get-Item -LiteralPath $path
        $sourceFile.CreationTime = [datetime]'2024-01-02T03:04:05'
        $sourceFile.LastWriteTime = [datetime]'2025-02-03T04:05:06'
        $creationTime = $sourceFile.CreationTime
        $lastWriteTime = $sourceFile.LastWriteTime
        $length = $sourceFile.Length
        $config = Get-Content "$scriptRoot\Example.json" -Raw | ConvertFrom-Json
        $task = [pscustomobject]@{
            TaskIndex = 0; ComputerName = 'localhost'; Name = '000123'; Path = $sourceFolder.FullName
            Recurse = $false; SharePoint = $config.Tasks[0].SharePoint
            AfterUpload = [pscustomobject]@{ RemoveFile = $true; CopyToFolder = (Join-Path $TestDrive ([guid]::NewGuid())) }
        }
        $results = @(& "$scriptRoot\Upload files.ps1" $task '[pscustomobject]@{ Id="uploaded"; WebUrl="https://example.com/report" }')
        Test-Path -LiteralPath $path | Should -BeFalse
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        Send-UploadReportHC $settings $results $systemErrors (Get-Date) 'Test'
        $systemErrors.Count | Should -Be 0
        $excel = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.xlsx'
        $overview = Import-Excel -Path $excel.FullName -WorksheetName 'Overview'
        $overview.FullName | Should -Be $path
        $overview.FileName | Should -Be '000123.txt'
        $overview.Name | Should -Be '000123'
        $overview.CreationTime | Should -Be $creationTime
        $overview.LastWriteTime | Should -Be $lastWriteTime
        $overview.Length | Should -Be $length
        $overview.Action | Should -Be 'Uploaded, Copied, Removed'
        $overview.DateTime | Should -BeOfType ([datetime])
        $package = Open-ExcelPackage -Path $excel.FullName
        try {
            $sheet = $package.Workbook.Worksheets['Overview']
            $sheet.Cells[1, 7].Value | Should -Be 'FullName'
            $sheet.Cells[1, 8].Value | Should -Be 'CreationTime'
            $sheet.Cells[1, 9].Value | Should -Be 'LastWriteTime'
            $sheet.Cells[1, 11].Value | Should -Be 'Action'
            $sheet.Cells[2, 8].Style.Numberformat.Format | Should -Not -Be 'General'
            $sheet.Cells[2, 9].Style.Numberformat.Format | Should -Not -Be 'General'
        }
        finally { $package.Dispose() }
    }

    It 'does not create JSON for file errors without system errors' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        $row.Uploaded = $false
        $row.Error = 'SharePoint file already exists'
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        $excel = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.xlsx'
        (Import-Excel $excel.FullName).Error | Should -Be $row.Error
        @(Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '*.json').Count | Should -Be 0
    }

    It 'logs Excel export failure as a system error' {
        $settings.SaveLogFiles.Where.Folder = Join-Path $TestDrive ([guid]::NewGuid())
        $settings.SendMail.When = 'Never'
        Mock Export-Excel { throw 'Workbook is locked' }
        Send-UploadReportHC $settings @($row) $systemErrors (Get-Date) 'Test'
        $json = Get-ChildItem $settings.SaveLogFiles.Where.Folder -Filter '* - System errors log.json'
        $errors = @(Get-Content $json.FullName -Raw | ConvertFrom-Json)
        $errors.Count | Should -Be 1
        $errors[0].Message | Should -BeLike '*Workbook is locked*'
    }
}