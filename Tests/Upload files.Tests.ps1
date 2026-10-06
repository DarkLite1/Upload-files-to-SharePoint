BeforeAll {
    $worker = Join-Path (Split-Path $PSScriptRoot) 'Upload files.ps1'
    $success = 'param($FilePath, $FolderPath) [pscustomobject]@{ Id = "item-id"; WebUrl = "https://example.com/$FolderPath" }'
    $getChildItemCommand = Get-Command Get-ChildItem -CommandType Cmdlet
}

Describe 'Upload worker post-upload safety' {
    BeforeEach {
        $source = New-Item -ItemType Directory -Path (Join-Path $TestDrive ([guid]::NewGuid()))
        $archive = Join-Path $TestDrive ([guid]::NewGuid())
        $sourcePath = Join-Path $source.FullName 'report[1].txt'
        Set-Content -LiteralPath $sourcePath -Value 'report'
        $task = [pscustomobject]@{
            TaskIndex = 0; ComputerName = 'localhost'; Name = 'Reports'; Path = $source.FullName
            Recurse = $false
            SharePoint = [pscustomobject]@{ SiteUrl = 'https://example.com'; DocumentLibraryName = 'Documents'; FolderPath = 'Reports'; ClientId = 'client'; TenantId = 'tenant'; CertificateThumbprint = 'cert' }
            AfterUpload = [pscustomobject]@{ RemoveFile = $false; CopyToFolder = $null }
        }
    }

    It 'keeps the source by default' {
        $result = & $worker $task $success
        $result.Uploaded | Should -BeTrue
        $result.Error | Should -BeNullOrEmpty
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
    }

    It 'does not copy or delete after a failed upload' {
        $task.AfterUpload.RemoveFile = $true
        $task.AfterUpload.CopyToFolder = $archive
        $result = & $worker $task 'throw "Upload failed"'
        $result.Error | Should -BeLike '*Upload failed*'
        $result.Uploaded | Should -BeFalse
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
        Test-Path -LiteralPath $archive | Should -BeFalse
    }

    It 'requires a confirmed item ID before removal' {
        $task.AfterUpload.RemoveFile = $true
        $result = & $worker $task '[pscustomobject]@{ WebUrl = "incomplete" }'
        $result.Error | Should -BeLike '*confirmed*'
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
    }

    It 'copies before removing the source' {
        $task.AfterUpload.RemoveFile = $true
        $task.AfterUpload.CopyToFolder = $archive
        $result = & $worker $task $success
        $result.Error | Should -BeNullOrEmpty
        $result.Removed | Should -BeTrue
        Get-Content -LiteralPath $result.CopiedTo | Should -Be 'report'
        Test-Path -LiteralPath $sourcePath | Should -BeFalse
    }

    It 'copies without removing when removal is disabled' {
        $task.AfterUpload.CopyToFolder = $archive
        $result = & $worker $task $success
        Test-Path -LiteralPath $result.CopiedTo | Should -BeTrue
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
    }

    It 'retains the source when the archive file already exists' {
        $task.AfterUpload.RemoveFile = $true
        $task.AfterUpload.CopyToFolder = $archive
        $null = New-Item -ItemType Directory $archive
        Set-Content -LiteralPath (Join-Path $archive 'report[1].txt') -Value 'older report'
        $result = & $worker $task $success
        $result.Uploaded | Should -BeTrue
        $result.Error | Should -Not -BeNullOrEmpty
        $result.Removed | Should -BeFalse
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
        Get-Content -LiteralPath (Join-Path $archive 'report[1].txt') | Should -Be 'older report'
    }

    It 'preserves relative subfolders in SharePoint and the archive' {
        $null = New-Item -ItemType Directory (Join-Path $source 'Sub')
        Set-Content (Join-Path $source 'Sub/nested.txt') 'nested'
        $task.Recurse = $true
        $task.AfterUpload.CopyToFolder = $archive
        $results = @(& $worker $task $success)
        $results.Count | Should -Be 2
        $nested = $results | Where-Object Destination -eq 'Reports/Sub/nested.txt'
        $nested.Uploaded | Should -BeTrue
        Get-Content -LiteralPath $nested.CopiedTo | Should -Be 'nested'
    }

    It 'does not enter subfolders when Recurse is false' {
        $null = New-Item -ItemType Directory (Join-Path $source 'Sub')
        Set-Content (Join-Path $source 'Sub/nested.txt') 'nested'
        @(& $worker $task $success).Count | Should -Be 1
    }

    It 'rejects an archive inside the source' {
        $task.AfterUpload.CopyToFolder = Join-Path $source 'Archive'
        { & $worker $task $success } | Should -Throw '*outside*'
    }

    It 'reports an empty folder without errors or uploads' {
        Remove-Item -LiteralPath $sourcePath
        @(& $worker $task $success).Count | Should -Be 0
    }

    It 'continues with other files after one upload fails' {
        Set-Content (Join-Path $source 'other.txt') 'other'
        $results = @(& $worker $task 'param($FilePath) if ($FilePath.EndsWith("other.txt")) { throw "Rejected" }; [pscustomobject]@{Id="ok"}')
        @($results | Where-Object Uploaded).Count | Should -Be 1
        @($results | Where-Object Error).Count | Should -Be 1
    }

    It 'skips <Attribute> files without copying or deleting them' -ForEach @(
        @{ Attribute = 'Hidden' }
        @{ Attribute = 'System' }
    ) {
        [IO.File]::SetAttributes($sourcePath, [IO.FileAttributes]$Attribute)
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @($Attribute) })
        $task.AfterUpload.RemoveFile = $true
        $task.AfterUpload.CopyToFolder = $archive
        @(& $worker $task 'throw "Excluded file reached uploader"').Count | Should -Be 0
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
        Test-Path -LiteralPath $archive | Should -BeFalse
    }

    It 'includes hidden and system files with omitted or empty exclusions' -ForEach @(
        @{ EmptyExclude = $false }
        @{ EmptyExclude = $true }
    ) {
        [IO.File]::SetAttributes($sourcePath, [IO.FileAttributes]'Hidden, System')
        if ($EmptyExclude) { $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @() }) }
        $result = & $worker $task $success
        $result.Uploaded | Should -BeTrue
        $result.Error | Should -BeNullOrEmpty
    }

    It 'does not exclude System files when only Hidden is selected' {
        [IO.File]::SetAttributes($sourcePath, [IO.FileAttributes]::System)
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden') })
        (& $worker $task $success).Uploaded | Should -BeTrue
    }

    It 'skips files matching either attribute but uploads ordinary files' {
        $hiddenPath = Join-Path $source 'hidden.txt'
        $systemPath = Join-Path $source 'system.txt'
        Set-Content $hiddenPath 'hidden'
        Set-Content $systemPath 'system'
        [IO.File]::SetAttributes($hiddenPath, [IO.FileAttributes]::Hidden)
        [IO.File]::SetAttributes($systemPath, [IO.FileAttributes]::System)
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden', 'System') })
        $results = @(& $worker $task $success)
        $results.Count | Should -Be 1
        $results[0].Source | Should -Be $sourcePath
    }

    It 'does not enter <Attribute> folders or upload their ordinary children' -ForEach @(
        @{ Attribute = 'Hidden' }
        @{ Attribute = 'System' }
    ) {
        $excluded = New-Item -ItemType Directory (Join-Path $source 'Excluded')
        Set-Content (Join-Path $excluded 'ordinary.txt') 'protected'
        $excluded.Attributes = $excluded.Attributes -bor [IO.FileAttributes]$Attribute
        $included = New-Item -ItemType Directory (Join-Path $source 'Included')
        Set-Content (Join-Path $included 'ordinary.txt') 'upload'
        $task.Recurse = $true
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @($Attribute) })
        Mock Get-ChildItem { & $getChildItemCommand -LiteralPath $LiteralPath -Force }
        Mock Get-ChildItem { throw 'Excluded directory was entered' } -ParameterFilter { $LiteralPath -eq $excluded.FullName }
        $results = @(& $worker $task $success)
        $results.Count | Should -Be 2
        @($results | Where-Object Error).Count | Should -Be 0
        $results.Destination | Should -Contain 'Reports/Included/ordinary.txt'
        Should -Invoke Get-ChildItem -Times 0 -ParameterFilter { $LiteralPath -eq $excluded.FullName }
    }

    It 'skips an attribute-excluded source root' {
        $source.Attributes = $source.Attributes -bor [IO.FileAttributes]::Hidden
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden') })
        Mock Get-ChildItem { throw 'Excluded root was entered' }
        @(& $worker $task $success).Count | Should -Be 0
        Should -Invoke Get-ChildItem -Times 0
    }

    It 'honors Recurse false with exclusions enabled' {
        $null = New-Item -ItemType Directory (Join-Path $source 'Sub')
        Set-Content (Join-Path $source 'Sub/nested.txt') 'nested'
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden') })
        @(& $worker $task $success).Count | Should -Be 1
    }

    It 'reports genuine read errors while exclusions are enabled' {
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden') })
        Mock Get-ChildItem { Write-Error 'Access denied' -TargetObject $LiteralPath -ErrorAction Continue }
        $results = @(& $worker $task $success)
        $results.Count | Should -Be 1
        $results[0].Error | Should -BeLike '*Access denied*'
        $results[0].Source | Should -Be $source.FullName
    }

    It 'retains an uploaded file that becomes excluded before post-upload actions' {
        $task | Add-Member Exclude ([pscustomobject]@{ Attributes = @('Hidden') })
        $task.AfterUpload.RemoveFile = $true
        $task.AfterUpload.CopyToFolder = $archive
        $result = & $worker $task 'param($FilePath) [IO.File]::SetAttributes($FilePath, [IO.FileAttributes]::Hidden); [pscustomobject]@{ Id = "uploaded" }'
        $result.Uploaded | Should -BeTrue
        $result.Removed | Should -BeFalse
        $result.CopiedTo | Should -BeNullOrEmpty
        Test-Path -LiteralPath $sourcePath | Should -BeTrue
    }
}