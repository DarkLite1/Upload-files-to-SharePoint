#Requires -Version 7

<#
.SYNOPSIS
    Upload the files below one source folder and apply successful-upload actions.
.DESCRIPTION
    Internal worker used locally or through PowerShell remoting by Main.ps1.
    UploadScript contains the bundled single-file uploader, so remote computers
    do not need a copy of this script folder. They do need the Graph module.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)][PSCustomObject]$Task,
    [Parameter(Mandatory)][String]$UploadScript
)

$ErrorActionPreference = 'Stop'
$uploader = [scriptblock]::Create($UploadScript)
$root = Get-Item -LiteralPath $Task.Path -Force
if (-not $root.PSIsContainer -or ($root.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw "Source '$($Task.Path)' must be a folder, not a link."
}
$rootPath = $root.FullName.TrimEnd('\', '/')
$copyRoot = $Task.AfterUpload.CopyToFolder
if ($copyRoot) {
    $copyRoot = [IO.Path]::GetFullPath($copyRoot).TrimEnd('\', '/')
    if ($copyRoot -eq $rootPath -or $copyRoot.StartsWith("$rootPath\", [StringComparison]::OrdinalIgnoreCase)) {
        throw "CopyToFolder '$copyRoot' must be outside the source folder."
    }
}

$enumerationErrors = @()
$files = @(Get-ChildItem -LiteralPath $root.FullName -File -Force -Recurse:([bool]$Task.Recurse) -ErrorAction SilentlyContinue -ErrorVariable enumerationErrors)
foreach ($failure in $enumerationErrors) {
    [PSCustomObject]@{
        TaskIndex = $Task.TaskIndex; ComputerName = $Task.ComputerName; Name = $Task.Name
        Source = "$($failure.TargetObject)"; Destination = ''; WebUrl = ''
        Uploaded = $false; CopiedTo = ''; Removed = $false; Error = "$failure"
    }
}

foreach ($sourceFile in $files) {
    $result = [PSCustomObject]@{
        TaskIndex = $Task.TaskIndex; ComputerName = $Task.ComputerName; Name = $Task.Name
        Source = $sourceFile.FullName; Destination = ''; WebUrl = ''
        Uploaded = $false; CopiedTo = ''; Removed = $false; Error = ''
    }
    $sourceLock = $null
    try {
        if ($sourceFile.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw 'Symbolic links and other reparse-point files are not uploaded.'
        }
        $relativePath = [IO.Path]::GetRelativePath($root.FullName, $sourceFile.FullName)
        $relativeFolder = [IO.Path]::GetDirectoryName($relativePath)
        $targetFolder = (@($Task.SharePoint.FolderPath, $relativeFolder) | Where-Object { $_ }) -join '/'
        $targetFolder = $targetFolder.Replace('\', '/').Trim('/')
        $result.Destination = (@($targetFolder, $sourceFile.Name) | Where-Object { $_ }) -join '/'
        $uploadParams = @{
            FilePath = $sourceFile.FullName
            SiteUrl = $Task.SharePoint.SiteUrl
            DocumentLibraryName = $Task.SharePoint.DocumentLibraryName
            FolderPath = $targetFolder
            ClientId = $Task.SharePoint.ClientId
            TenantId = $Task.SharePoint.TenantId
            CertificateThumbprint = $Task.SharePoint.CertificateThumbprint
        }
        foreach ($property in 'MaxRetries', 'ChunkSizeMB') {
            if ($null -ne $Task.SharePoint.$property) { $uploadParams[$property] = $Task.SharePoint.$property }
        }
        $sourceLock = [IO.File]::Open($sourceFile.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $uploaded = & $uploader @uploadParams
        if (-not $uploaded.Id) { throw 'Upload did not return a confirmed SharePoint item ID.' }
        $result.Uploaded = $true
        $result.WebUrl = $uploaded.WebUrl
        if ($copyRoot) {
            $copyPath = Join-Path $copyRoot $relativePath
            $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copyPath))
            [IO.File]::Copy($sourceFile.FullName, $copyPath, $false)
            $result.CopiedTo = $copyPath
        }
        $sourceFile.Refresh()
        $uploadedLength = $sourceFile.Length
        $uploadedWriteTime = $sourceFile.LastWriteTimeUtc
        $sourceLock.Dispose()
        $sourceLock = $null
        if ($Task.AfterUpload.RemoveFile) {
            $sourceFile.Refresh()
            if (-not $sourceFile.Exists -or $sourceFile.Length -ne $uploadedLength -or $sourceFile.LastWriteTimeUtc -ne $uploadedWriteTime) {
                throw 'Source changed after upload; it was not removed.'
            }
            Remove-Item -LiteralPath $sourceFile.FullName -Force
            $result.Removed = $true
        }
    }
    catch { $result.Error = "$($_.Exception.Message)" }
    finally { if ($sourceLock) { $sourceLock.Dispose() } }
    $result
}