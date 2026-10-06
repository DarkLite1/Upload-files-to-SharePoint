#Requires -Version 7

<#
.SYNOPSIS
    Upload the files below one source folder and apply successful-upload actions.
.DESCRIPTION
    Internal worker used locally or through PowerShell remoting by Main.ps1.
    UploadScript contains the bundled single-file uploader, so remote computers
    do not need a copy of this script folder. They do need the Graph module.
    Task.Exclude groups Attributes (Hidden/System), Folders (subfolder trees),
    and Files (literal file paths). Omitted or empty lists exclude nothing.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory)][PSCustomObject]$Task,
    [Parameter(Mandatory)][String]$UploadScript
)

$ErrorActionPreference = 'Stop'
$uploader = [scriptblock]::Create($UploadScript)
$excludedAttributeMask = [IO.FileAttributes]0
foreach ($attribute in $Task.Exclude.Attributes) {
    $excludedAttributeMask = $excludedAttributeMask -bor [IO.FileAttributes]$attribute
}
$excludedFolders = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$excludedFiles = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($path in $Task.Exclude.Folders) {
    $null = $excludedFolders.Add([IO.Path]::GetFullPath($path.Replace('/', '\')).TrimEnd('\'))
}
foreach ($path in $Task.Exclude.Files) {
    $null = $excludedFiles.Add([IO.Path]::GetFullPath($path.Replace('/', '\')).TrimEnd('\'))
}
$root = Get-Item -LiteralPath $Task.Path -Force
if (-not $root.PSIsContainer -or ($root.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
    throw "Source '$($Task.Path)' must be a folder, not a link."
}
if ($root.Attributes -band $excludedAttributeMask) { return }
$rootPath = $root.FullName.TrimEnd('\', '/')
$copyRoot = $Task.AfterUpload.CopyToFolder
if ($copyRoot) {
    $copyRoot = [IO.Path]::GetFullPath($copyRoot).TrimEnd('\', '/')
    if ($copyRoot -eq $rootPath -or $copyRoot.StartsWith("$rootPath\", [StringComparison]::OrdinalIgnoreCase)) {
        throw "CopyToFolder '$copyRoot' must be outside the source folder."
    }
}

$enumerationErrors = [Collections.Generic.List[object]]::new()
$files = @(
    if ($excludedAttributeMask -or $excludedFolders.Count) {
        $pending = [Collections.Generic.Stack[string]]::new()
        $pending.Push($root.FullName)
        while ($pending.Count) {
            $directoryPath = $pending.Pop()
            if ($excludedFolders.Contains($directoryPath.TrimEnd('\'))) { continue }
            try {
                $attributes = [IO.File]::GetAttributes($directoryPath)
                if ($attributes -band ($excludedAttributeMask -bor [IO.FileAttributes]::ReparsePoint)) { continue }
            }
            catch {
                $enumerationErrors.Add([Management.Automation.ErrorRecord]::new(
                    $_.Exception, $_.FullyQualifiedErrorId, $_.CategoryInfo.Category, $directoryPath
                ))
                continue
            }
            $readErrors = @()
            Get-ChildItem -LiteralPath $directoryPath -Force -ErrorAction SilentlyContinue -ErrorVariable readErrors |
                ForEach-Object {
                    if ($_.Attributes -band $excludedAttributeMask) { return }
                    if ($_.PSIsContainer) {
                        if ($Task.Recurse -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -and -not $excludedFolders.Contains($_.FullName.TrimEnd('\'))) {
                            $pending.Push($_.FullName)
                        }
                    }
                    else { $_ }
                }
            foreach ($failure in $readErrors) { $enumerationErrors.Add($failure) }
        }
    }
    else {
        $readErrors = @()
        Get-ChildItem -LiteralPath $root.FullName -File -Force -Recurse:([bool]$Task.Recurse) -ErrorAction SilentlyContinue -ErrorVariable readErrors
        foreach ($failure in $readErrors) { $enumerationErrors.Add($failure) }
    }
)
foreach ($failure in $enumerationErrors) {
    [PSCustomObject]@{
        TaskIndex = $Task.TaskIndex; ComputerName = $Task.ComputerName; Name = $Task.Name
        DateTime = Get-Date; FullName = "$($failure.TargetObject)"; FileName = ''; Type = 'Folder'
        CreationTime = $null; LastWriteTime = $null; Length = $null; Action = ''
        Source = "$($failure.TargetObject)"; Destination = ''; WebUrl = ''
        Uploaded = $false; CopiedTo = ''; Removed = $false; Error = "$failure"
        Stage = 'Read folder'; TargetObject = "$($failure.TargetObject)"
        FullyQualifiedErrorId = $failure.FullyQualifiedErrorId
        ExceptionType = $failure.Exception.GetType().FullName
        ScriptStackTrace = $failure.ScriptStackTrace; PositionMessage = $failure.InvocationInfo.PositionMessage
    }
}

foreach ($sourceFile in $files) {
    if ($excludedFiles.Contains($sourceFile.FullName)) { continue }
    $result = [PSCustomObject]@{
        TaskIndex = $Task.TaskIndex; ComputerName = $Task.ComputerName; Name = $Task.Name
        DateTime = Get-Date; FullName = $sourceFile.FullName; FileName = $sourceFile.Name; Type = 'File'
        CreationTime = $null; LastWriteTime = $null; Length = $null; Action = ''
        Source = $sourceFile.FullName; Destination = ''; WebUrl = ''
        Uploaded = $false; CopiedTo = ''; Removed = $false; Error = ''
        Stage = 'Read file'; TargetObject = ''; FullyQualifiedErrorId = ''; ExceptionType = ''
        ScriptStackTrace = ''; PositionMessage = ''
    }
    $sourceLock = $null
    try {
        $sourceFile.Refresh()
        if (-not $sourceFile.Exists) { throw "Source file '$($sourceFile.FullName)' is no longer accessible." }
        if ($sourceFile.Attributes -band $excludedAttributeMask) { continue }
        $result.CreationTime = $sourceFile.CreationTime
        $result.LastWriteTime = $sourceFile.LastWriteTime
        $result.Length = $sourceFile.Length
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
            Overwrite = [bool]$Task.SharePoint.Overwrite
        }
        foreach ($property in 'MaxRetries', 'ChunkSizeMB') {
            if ($null -ne $Task.SharePoint.$property) { $uploadParams[$property] = $Task.SharePoint.$property }
        }
        $sourceLock = [IO.File]::Open($sourceFile.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $result.Stage = 'Upload file'
        $uploaded = & $uploader @uploadParams
        if (-not $uploaded.Id) { throw 'Upload did not return a confirmed SharePoint item ID.' }
        $result.Uploaded = $true
        $result.Action = 'Uploaded'
        $result.WebUrl = $uploaded.WebUrl
        $sourceFile.Refresh()
        $isExcluded = [bool]($sourceFile.Attributes -band $excludedAttributeMask)
        if ($copyRoot -and -not $isExcluded) {
            $result.Stage = 'Copy file'
            $copyPath = Join-Path $copyRoot $relativePath
            $null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copyPath))
            [IO.File]::Copy($sourceFile.FullName, $copyPath, $false)
            $result.CopiedTo = $copyPath
            $result.Action += ', Copied'
        }
        $sourceFile.Refresh()
        $uploadedLength = $sourceFile.Length
        $uploadedWriteTime = $sourceFile.LastWriteTimeUtc
        $sourceLock.Dispose()
        $sourceLock = $null
        if ($Task.AfterUpload.RemoveFile -and -not $isExcluded) {
            $result.Stage = 'Remove file'
            $sourceFile.Refresh()
            if (-not $sourceFile.Exists -or $sourceFile.Length -ne $uploadedLength -or $sourceFile.LastWriteTimeUtc -ne $uploadedWriteTime) {
                throw 'Source changed after upload; it was not removed.'
            }
            if (-not ($sourceFile.Attributes -band $excludedAttributeMask)) {
                Remove-Item -LiteralPath $sourceFile.FullName -Force
                $result.Removed = $true
                $result.Action += ', Removed'
            }
        }
    }
    catch {
        $result.Error = "$($_.Exception.Message)"
        $result.TargetObject = "$($_.TargetObject)"
        $result.FullyQualifiedErrorId = $_.FullyQualifiedErrorId
        $result.ExceptionType = $_.Exception.GetType().FullName
        $result.ScriptStackTrace = $_.ScriptStackTrace
        $result.PositionMessage = $_.InvocationInfo.PositionMessage
    }
    finally { if ($sourceLock) { $sourceLock.Dispose() } }
    $result
}