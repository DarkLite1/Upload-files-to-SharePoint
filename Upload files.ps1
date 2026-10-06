#Requires -Version 7

<#
.SYNOPSIS
    Upload the files below one source folder and apply successful-upload actions.
.DESCRIPTION
    Internal worker used locally or through PowerShell remoting by Main.ps1.
    UploadScript contains the bundled single-file uploader, so remote computers
    do not need a copy of this script folder. They do need the Graph module.
    Task.Exclude.Attributes optionally skips Hidden/System files and entire
    matching folder trees. Omitted or empty exclusions include all attributes.
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
    if ($excludedAttributeMask) {
        $pending = [Collections.Generic.Stack[string]]::new()
        $pending.Push($root.FullName)
        while ($pending.Count) {
            $directoryPath = $pending.Pop()
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
                        if ($Task.Recurse -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
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
        $sourceFile.Refresh()
        if (-not $sourceFile.Exists) { throw "Source file '$($sourceFile.FullName)' is no longer accessible." }
        if ($sourceFile.Attributes -band $excludedAttributeMask) { continue }
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
        $uploaded = & $uploader @uploadParams
        if (-not $uploaded.Id) { throw 'Upload did not return a confirmed SharePoint item ID.' }
        $result.Uploaded = $true
        $result.WebUrl = $uploaded.WebUrl
        $sourceFile.Refresh()
        $isExcluded = [bool]($sourceFile.Attributes -band $excludedAttributeMask)
        if ($copyRoot -and -not $isExcluded) {
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
        if ($Task.AfterUpload.RemoveFile -and -not $isExcluded) {
            $sourceFile.Refresh()
            if (-not $sourceFile.Exists -or $sourceFile.Length -ne $uploadedLength -or $sourceFile.LastWriteTimeUtc -ne $uploadedWriteTime) {
                throw 'Source changed after upload; it was not removed.'
            }
            if (-not ($sourceFile.Attributes -band $excludedAttributeMask)) {
                Remove-Item -LiteralPath $sourceFile.FullName -Force
                $result.Removed = $true
            }
        }
    }
    catch { $result.Error = "$($_.Exception.Message)" }
    finally { if ($sourceLock) { $sourceLock.Dispose() } }
    $result
}