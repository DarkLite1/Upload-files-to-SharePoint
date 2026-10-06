#Requires -Version 7
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Upload selected files and folder contents to SharePoint document libraries.
.DESCRIPTION
    Tasks select sources through Upload.Folders and Upload.Files.
    AfterUpload can copy and/or remove a source
    file, only after its upload succeeds. With both enabled, copying must
    succeed before removal. Folder jobs preserve relative subfolders when
    recursing; explicit files can override the destination via SharePointFileName.
    Reports errors and exits with code 1 if any upload or reporting step fails.
.PARAMETER ConfigurationJsonFile
    JSON configuration path, relative to the current working directory or absolute.
.EXAMPLE
    .\Main.ps1 -ConfigurationJsonFile .\Example.json
#>
[CmdletBinding()]
param ([Parameter(Mandatory)][string]$ConfigurationJsonFile)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Functions.ps1"
$started = Get-Date
$systemErrors = [Collections.Generic.List[object]]::new()
$results = @()
$settings = $null
$configurationName = [IO.Path]::GetFileNameWithoutExtension($ConfigurationJsonFile)
try {
    $configuration = Get-Content -LiteralPath $ConfigurationJsonFile -Raw -Encoding utf8 | ConvertFrom-Json
    $settings = $configuration.Settings
    $tasks = @(ConvertTo-UploadConfigurationHC $configuration)
    $uploadScript = Get-Content -LiteralPath "$PSScriptRoot\UploadToSharePoint.ps1" -Raw
    $workerPath = "$PSScriptRoot\Upload files.ps1"
    $sessionConfiguration = if ($configuration.PSSessionConfiguration) { $configuration.PSSessionConfiguration } else { 'PowerShell.7' }
    $workers = @($tasks | Group-Object Target | ForEach-Object {
        $queue = [Collections.Concurrent.ConcurrentQueue[object]]::new()
        foreach ($task in $_.Group) { $queue.Enqueue($task) }
        $workerCount = [math]::Min([int]$configuration.MaxConcurrent.JobsPerComputer, $queue.Count)
        for ($workerIndex = 0; $workerIndex -lt $workerCount; $workerIndex++) {
            [pscustomobject]@{ Queue = $queue }
        }
    })
    $execute = {
        param ($Worker, $WorkerPath, $UploadScript, $SessionConfiguration)
        $ErrorActionPreference = 'Stop'
        $task = $null
        while ($Worker.Queue.TryDequeue([ref]$task)) {
            $session = $null
            try {
                if ($task.ComputerName -eq $env:COMPUTERNAME) {
                    $stage = 'Run local worker'
                    & $WorkerPath -Task $task -UploadScript $UploadScript
                }
                else {
                    $stage = 'Open remote session'
                    $session = New-PSSession -ComputerName $task.ComputerName -ConfigurationName $SessionConfiguration -ErrorAction Stop
                    $stage = 'Run remote worker'
                    Invoke-Command -Session $session -FilePath $WorkerPath -ArgumentList $task, $UploadScript -ErrorAction Stop
                }
            }
            catch {
                $exception = if ($_.Exception.SerializedRemoteException) { $_.Exception.SerializedRemoteException } else { $_.Exception }
                $positionMessage = if ($_.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
                    $_.Exception.SerializedRemoteInvocationInfo.PositionMessage
                }
                else { $_.InvocationInfo.PositionMessage }
                [pscustomobject]@{
                    TaskIndex = $task.TaskIndex; ComputerName = $task.ComputerName; Name = $task.Name
                    DateTime = Get-Date; FullName = $task.Path; FileName = ''; Type = 'Task'
                    CreationTime = $null; LastWriteTime = $null; Length = $null; Action = ''
                    Source = $task.Path; Destination = ''; WebUrl = ''; Uploaded = $false; CopiedTo = ''; Removed = $false
                    Error = "$($_.Exception.Message)"
                    Stage = $stage; TargetObject = "$($_.TargetObject)"; FullyQualifiedErrorId = $_.FullyQualifiedErrorId
                    ExceptionType = $exception.PSTypeNames[0] -replace '^Deserialized\.', ''
                    ScriptStackTrace = $_.ScriptStackTrace; PositionMessage = $positionMessage
                }
            }
            finally {
                if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
            }
        }
    }
    $results = @(Invoke-WithOptionalParallelismHC -InputObject $workers -ScriptBlock $execute -ThrottleLimit $configuration.MaxConcurrent.JobsTotal -ArgumentList $workerPath, $uploadScript, $sessionConfiguration)
}
catch { $systemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "$($_.Exception.Message)" }) }

try { Send-UploadReportHC -Settings $settings -Results $results -SystemErrors $systemErrors -Started $started -ConfigurationName $configurationName }
catch { $systemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Reporting failed: $_" }) }
$results
foreach ($failure in $systemErrors) { Write-Warning $failure.Message }
if ($systemErrors.Count -or @($results | Where-Object Error).Count) {
    Write-Warning 'Exit script with error code 1. Inspect Error in the returned results and the log files.'
    exit 1
}