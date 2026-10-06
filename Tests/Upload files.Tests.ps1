BeforeAll {
    $worker = Join-Path (Split-Path $PSScriptRoot) 'Upload files.ps1'
    $success = 'param($FilePath, $FolderPath) [pscustomobject]@{ Id = "item-id"; WebUrl = "https://example.com/$FolderPath" }'
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
}