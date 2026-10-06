function Get-LogFolderHC {
    <#
    .SYNOPSIS
        Create the log folder when it doesn't exist and return its full
        path. Relative paths are relative to the script folder.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
        $fullPath = $Path
    }
    else {
        $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
    }

    if (-not (Test-Path -Path $fullPath -PathType Container)) {
        try {
            Write-Verbose "Create log folder '$fullPath'"
            $null = New-Item -Path $fullPath -ItemType Directory -Force -ErrorAction Stop
        }
        catch {
            throw "Failed creating log folder '$fullPath': $_"
        }
    }

    (Resolve-Path $fullPath).ProviderPath
}

function Get-StringValueHC {
    <#
    .SYNOPSIS
        Retrieve a string from the environment variables or a regular
        string.

    .DESCRIPTION
        When the value starts with 'ENV:' the value of that environment
        variable is returned, otherwise the value itself.

    .EXAMPLE
        Get-StringValueHC -Name 'ENV:passwordVariable'

        # Output: the value of $ENV:passwordVariable or an error when the
        # variable does not exist
    #>
    param (
        [String]$Name
    )

    if (-not $Name) {
        return $null
    }
    elseif (
        $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        $envVariableName = $Name.Substring(4).Trim()
        $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
        if ($envStringValue) {
            return $envStringValue.Value
        }
        else {
            throw "Environment variable '$envVariableName' not found."
        }
    }
    else {
        return $Name
    }
}

function Invoke-WithOptionalParallelismHC {
    <#
    .SYNOPSIS
        Run a scriptblock for each input object, sequentially or in
        parallel.

    .DESCRIPTION
        With a ThrottleLimit of 1 or less the scriptblock runs in a plain
        foreach loop on the main thread. Otherwise it runs with
        ForEach-Object -Parallel.

        The input object is passed as the first positional argument,
        followed by the values in ArgumentList.

        The scriptblock is rehydrated from its text inside each parallel
        runspace, so '$using:' does not work inside it. Pass everything it
        needs through the input object (DTO) or ArgumentList, and return
        results instead of changing shared objects.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$InputObject,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [Parameter(Mandatory)]
        [int]$ThrottleLimit,
        [object[]]$ArgumentList = @()
    )

    if ($ThrottleLimit -le 1) {
        foreach ($item in $InputObject) {
            & $ScriptBlock $item @ArgumentList
        }
    }
    else {
        $scriptBlockString = $ScriptBlock.ToString()

        $InputObject | ForEach-Object -Parallel {
            $rehydratedBlock = [scriptblock]::Create($using:scriptBlockString)
            $splatArgs = $using:ArgumentList
            & $rehydratedBlock $_ @splatArgs
        } -ThrottleLimit $ThrottleLimit
    }
}

function Send-MailKitMessageHC {
    <#
    .SYNOPSIS
        Send an email using MailKit and MimeKit assemblies.

    .DESCRIPTION
        Requires the assemblies to be installed:

        $params = @{
            Source           = 'https://www.nuget.org/api/v2'
            SkipDependencies = $true
            Scope            = 'AllUsers'
        }
        Install-Package @params -Name 'MailKit'
        Install-Package @params -Name 'MimeKit'

        Supply paths to assemblies compatible with the installed PowerShell
        runtime. At least one To or Bcc recipient is required. A connection,
        authentication or sending failure throws an error.

    .PARAMETER MailKitAssemblyPath
        Path to MailKit.dll. Used if MailKit is not already loaded.

    .PARAMETER MimeKitAssemblyPath
        Path to MimeKit.dll. Used if MimeKit is not already loaded.

    .PARAMETER SmtpServerName
        SMTP server hostname or IP address.

    .PARAMETER SmtpPort
        SMTP port: 25, 465, 587 or 2525.

    .PARAMETER Body
        Email body as HTML.

    .PARAMETER Subject
        Complete subject line. This helper does not add counts or a prefix.

    .PARAMETER From
        Sender email address.

    .PARAMETER FromDisplayName
        Optional friendly name for the sender.

    .PARAMETER To
        Recipient addresses. To or Bcc must contain at least one address.

    .PARAMETER Bcc
        Blind-copy recipient addresses.

    .PARAMETER MaxAttachmentSize
        Combined source-file size limit in bytes. Defaults to 20 MB. At or
        above the limit, no attachments are added and a notice is appended
        to the body. This is not the size of the encoded email.

    .PARAMETER SmtpConnectionType
        MailKit SecureSocketOptions value. Defaults to None (no encryption).
        Use the mode required by your SMTP server, such as StartTls or
        SslOnConnect.

    .PARAMETER Priority
        Normal (default), Low or High. Sets the X-Priority header.

    .PARAMETER Attachments
        File paths to attach. Duplicates are removed. Missing files, folders
        and attachment failures produce warnings; sending can still proceed.

    .PARAMETER Credential
        Optional SMTP credential. If omitted, authentication is not attempted.
    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [string]$MailKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$MimeKitAssemblyPath,
        [parameter(Mandatory)]
        [string]$SmtpServerName,
        [parameter(Mandatory)]
        [ValidateSet(25, 465, 587, 2525)]
        [int]$SmtpPort,
        [parameter(Mandatory)]
        [string]$Body,
        [parameter(Mandatory)]
        [string]$Subject,
        [parameter(Mandatory)]
        [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
        [string]$From,
        [string]$FromDisplayName,
        [string[]]$To,
        [string[]]$Bcc,
        [int]$MaxAttachmentSize = 20MB,
        [ValidateSet(
            'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
        )]
        [string]$SmtpConnectionType = 'None',
        [ValidateSet('Normal', 'Low', 'High')]
        [string]$Priority = 'Normal',
        [string[]]$Attachments,
        [PSCredential]$Credential
    )

    begin {
        function Test-IsAssemblyLoaded {
            <#
            .SYNOPSIS
                Check whether an assembly with the given name is already
                loaded in this process, regardless of its version.
            #>
            param (
                [String]$Name
            )
            foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                if ($assembly.FullName -like "$Name, Version=*") {
                    return $true
                }
            }
            return $false
        }

        function Add-Attachments {
            <#
            .SYNOPSIS
                Add readable files to a MIME body within the attachment limit.

            .DESCRIPTION
                Uses the enclosing function's MaxAttachmentSize. If the
                combined size reaches the limit, returns an object with
                AttachmentLimitExceededMessage and adds no attachments.
                Other attachment problems are warnings.
            #>
            param (
                [string[]]$Attachments,
                [MimeKit.Multipart]$BodyMultiPart
            )

            $attachmentList = New-Object System.Collections.ArrayList($null)

            foreach (
                $attachmentPath in
                $Attachments | Sort-Object -Unique
            ) {
                try {
                    try {
                        $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                        if ($attachmentItem.PSIsContainer) {
                            Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                            continue
                        }
                    }
                    catch {
                        Write-Warning "Attachment '$attachmentPath' not found"
                        continue
                    }

                    $totalSizeAttachments += $attachmentItem.Length

                    $null = $attachmentList.Add($attachmentItem)

                    if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                        $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                        ([math]::Round(($MaxAttachmentSize / 1MB))),
                        ([math]::Round(($totalSizeAttachments / 1MB), 2))

                        Write-Warning $M

                        return [PSCustomObject]@{
                            AttachmentLimitExceededMessage = $M
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentPath': $_"
                }
            }

            foreach (
                $attachmentItem in
                $attachmentList
            ) {
                try {
                    Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                    $attachment = New-Object MimeKit.MimePart

                    $memoryStream = New-Object System.IO.MemoryStream

                    try {
                        $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                        $fileStream.CopyTo($memoryStream)
                    }
                    finally {
                        if ($fileStream) {
                            $fileStream.Dispose()
                        }
                    }

                    $memoryStream.Position = 0

                    $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                    $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                    $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                    $attachment.FileName = $attachmentItem.Name

                    $bodyMultiPart.Add($attachment)
                }
                catch {
                    Write-Warning "Failed to add attachment '$attachmentItem': $_"
                }
            }
        }

        try {
            if (-not ($To -or $Bcc)) {
                throw "Either 'To' to 'Bcc' is required for sending emails"
            }

            foreach ($email in $To) {
                if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    throw "To email address '$email' not valid."
                }
            }

            foreach ($email in $Bcc) {
                if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                    throw "Bcc email address '$email' not valid."
                }
            }

            if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                try {
                    Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                    Add-Type -Path $MimeKitAssemblyPath
                }
                catch {
                    throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                }
            }

            if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                try {
                    Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                    Add-Type -Path $MailKitAssemblyPath
                }
                catch {
                    throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                }
            }
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
    }

    process {
        try {
            $message = New-Object -TypeName 'MimeKit.MimeMessage'

            $bodyPart = New-Object MimeKit.TextPart('html')
            $bodyPart.Text = $Body

            $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
            $bodyMultiPart.Add($bodyPart)

            if ($Attachments) {
                $params = @{
                    Attachments   = $Attachments
                    BodyMultiPart = $bodyMultiPart
                }
                $addAttachments = Add-Attachments @params

                if ($addAttachments.AttachmentLimitExceededMessage) {
                    $bodyPart.Text += '<p><i>{0}</i></p>' -f
                    $addAttachments.AttachmentLimitExceededMessage
                }
            }

            $message.Body = $bodyMultiPart

            $fromAddress = New-Object MimeKit.MailboxAddress(
                $FromDisplayName, $From
            )
            $message.From.Add($fromAddress)

            foreach ($email in $To) {
                $message.To.Add($email)
            }

            foreach ($email in $Bcc) {
                $message.Bcc.Add($email)
            }

            $message.Subject = $Subject

            switch ($Priority) {
                'Low' {
                    $message.Headers.Add('X-Priority', '5 (Lowest)')
                    break
                }
                'Normal' {
                    $message.Headers.Add('X-Priority', '3 (Normal)')
                    break
                }
                'High' {
                    $message.Headers.Add('X-Priority', '1 (Highest)')
                    break
                }
                default {
                    throw "Priority type '$_' not supported"
                }
            }

            $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

            try {
                $smtp.Connect(
                    $SmtpServerName, $SmtpPort,
                    [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                )
            }
            catch {
                throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
            }

            if ($Credential) {
                try {
                    $smtp.Authenticate(
                        $Credential.UserName,
                        $Credential.GetNetworkCredential().Password
                    )
                }
                catch {
                    throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                }
            }

            Write-Verbose "Send mail to '$To' with subject '$Subject'"

            $null = $smtp.Send($message)
        }
        catch {
            throw "Failed to send email to '$To': $_"
        }
        finally {
            if ($smtp) {
                $smtp.Disconnect($true)
                $smtp.Dispose()
            }
            if ($message) {
                $message.Dispose()
            }
        }
    }
}

function Write-EventsToEventLogHC {
    <#
    .SYNOPSIS
        Write events to the event log.

    .DESCRIPTION
        Custom EventID's based on the PowerShell streams:
        100 Script started, 4 Verbose, 1 Output, 3 Warning, 2 Error,
        199 Script ended.

        All properties of an event that are not 'EntryType' or 'EventID'
        are used to create the message.
    #>

    [CmdLetBinding()]
    param (
        [Parameter(Mandatory)]
        [String]$Source,
        [Parameter(Mandatory)]
        [String]$LogName,
        [PSCustomObject[]]$Events
    )

    try {
        if ([System.Diagnostics.EventLog]::SourceExists($Source)) {
            $existingLogName = [System.Diagnostics.EventLog]::LogNameFromSourceName($Source, '.')

            if ($existingLogName -ne $LogName) {
                throw "The event log source '$Source' is already registered with event log name '$existingLogName', it cannot be used with log name '$LogName'."
            }
        }
        else {
            Write-Verbose "Create event log source '$Source' with log name '$LogName'"

            New-EventLog -LogName $LogName -Source $Source -EA Stop
        }

        foreach ($eventItem in $Events) {
            $params = @{
                LogName     = $LogName
                Source      = $Source
                EntryType   = $eventItem.EntryType
                EventID     = $eventItem.EventID
                Message     = ''
                ErrorAction = 'Stop'
            }

            if (-not $params.EntryType) {
                $params.EntryType = 'Information'
            }
            if (-not $params.EventID) {
                $params.EventID = 4
            }

            foreach (
                $property in
                $eventItem.PSObject.Properties | Where-Object {
                    ($_.Name -ne 'EntryType') -and ($_.Name -ne 'EventID')
                }
            ) {
                $params.Message += "`n- $($property.Name) '$($property.Value)'"
            }

            Write-EventLog @params
        }
    }
    catch {
        throw "Failed to write to event log '$LogName' source '$Source': $_"
    }
}

function ConvertTo-UploadConfigurationHC {
    [CmdletBinding()]
    param ([Parameter(Mandatory)][PSCustomObject]$Configuration)

    $settings = $Configuration.Settings
    if (-not $settings.ScriptName) { throw "Property 'Settings.ScriptName' not found" }
    if ($settings.ScriptName.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw 'Settings.ScriptName must be a valid file name.'
    }
    $mail = $settings.SendMail
    if ($mail.When -notin 'Never', 'Always', 'OnError', 'OnErrorOrAction') {
        throw 'Settings.SendMail.When must be Never, Always, OnError or OnErrorOrAction.'
    }
    if ($mail.When -ne 'Never') {
        foreach ($value in $mail.From, $mail.Smtp.ServerName, $mail.Smtp.Port, $mail.AssemblyPath.MailKit, $mail.AssemblyPath.MimeKit) {
            if (-not $value) { throw 'Settings.SendMail requires From, Smtp.ServerName, Smtp.Port and both AssemblyPath values.' }
        }
        if (-not ($mail.To -or $mail.Bcc)) { throw 'Settings.SendMail requires To or Bcc.' }
        if ((Get-StringValueHC "$($mail.Smtp.Port)") -notin 25, 465, 587, 2525) { throw 'Unsupported Settings.SendMail.Smtp.Port.' }
        if ($mail.Smtp.ConnectionType -and (Get-StringValueHC $mail.Smtp.ConnectionType) -notin 'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable') {
            throw 'Unsupported Settings.SendMail.Smtp.ConnectionType.'
        }
        if ([bool]$mail.Smtp.UserName -ne [bool]$mail.Smtp.Password) { throw 'Provide both SMTP UserName and Password, or neither.' }
    }
    if ($settings.SaveInEventLog.Save -isnot [bool]) { throw 'Settings.SaveInEventLog.Save must be true or false.' }
    if ($settings.SaveInEventLog.Save -and -not $settings.SaveInEventLog.LogName) { throw 'Settings.SaveInEventLog.LogName is required.' }
    $retention = $settings.SaveLogFiles.DeleteLogsAfterDays
    if ($null -ne $retention -and ("$retention" -notmatch '^\d+$' -or [decimal]$retention -gt [int]::MaxValue)) {
        throw 'Settings.SaveLogFiles.DeleteLogsAfterDays must be a nonnegative whole number.'
    }
    foreach ($property in 'JobsTotal', 'JobsPerComputer') {
        $value = $Configuration.MaxConcurrent.$property
        if ("$value" -notmatch '^\d+$' -or [decimal]$value -lt 1 -or [decimal]$value -gt [int]::MaxValue) {
            throw "MaxConcurrent.$property must be a positive whole number."
        }
    }
    foreach ($property in $Configuration.PSObject.Properties.Name) {
        if ($property -notin '?', 'Tasks', 'MaxConcurrent', 'Settings', 'PSSessionConfiguration') { throw "Unsupported property '$property'." }
    }
    if ($Configuration.Tasks -isnot [array] -or $Configuration.Tasks.Count -eq 0) { throw 'Tasks must be a nonempty array.' }
    $prepared = [Collections.Generic.List[object]]::new()
    $authentication = $null
    for ($taskIndex = 0; $taskIndex -lt $Configuration.Tasks.Count; $taskIndex++) {
        $task = $Configuration.Tasks[$taskIndex]
        $prefix = "Tasks[$taskIndex]"
        foreach ($property in $task.PSObject.Properties.Name) {
            if ($property -notin '?', 'ComputerName', 'Upload', 'Recurse', 'Exclude', 'SharePoint', 'AfterUpload') { throw "Unsupported property '$prefix.$property'." }
        }
        if ($task.Upload -isnot [pscustomobject]) { throw "$prefix.Upload must be an object with Folders and/or Files arrays." }
        foreach ($property in $task.Upload.PSObject.Properties.Name) {
            if ($property -notin '?', 'Folders', 'Files') { throw "Unsupported property '$prefix.Upload.$property'." }
        }
        $selections = @(
            foreach ($property in 'Folders', 'Files') {
                if ($task.Upload.PSObject.Properties.Name -notcontains $property) { continue }
                if ($task.Upload.$property -isnot [array]) { throw "$prefix.Upload.$property must be an array." }
                foreach ($entry in $task.Upload.$property) {
                    [pscustomobject]@{ Kind = $property; Entry = $entry }
                }
            }
        )
        if (-not $selections.Count) { throw "$prefix.Upload must select at least one folder or file." }
        $excludeAttributes = @()
        $excludedPaths = @{ Folders = @(); Files = @() }
        if ($task.PSObject.Properties.Name -contains 'Exclude') {
            if ($task.Exclude -isnot [pscustomobject]) { throw "$prefix.Exclude must be an object." }
            foreach ($property in $task.Exclude.PSObject.Properties.Name) {
                if ($property -notin '?', 'Attributes', 'Folders', 'Files') { throw "Unsupported property '$prefix.Exclude.$property'." }
            }
            if ($task.Exclude.PSObject.Properties.Name -contains 'Attributes') {
                if ($task.Exclude.Attributes -isnot [array]) { throw "$prefix.Exclude.Attributes must be an array containing Hidden and/or System." }
                foreach ($attribute in $task.Exclude.Attributes) {
                    if ($attribute -isnot [string] -or $attribute -notin 'Hidden', 'System') {
                        throw "$prefix.Exclude.Attributes supports only Hidden and System."
                    }
                }
                $excludeAttributes = @($task.Exclude.Attributes | Select-Object -Unique)
            }
            foreach ($property in 'Folders', 'Files') {
                if ($task.Exclude.PSObject.Properties.Name -notcontains $property) { continue }
                if ($task.Exclude.$property -isnot [array]) { throw "$prefix.Exclude.$property must be an array of paths." }
                $excludedPaths[$property] = @(
                    foreach ($excludedPath in $task.Exclude.$property) {
                        if ($excludedPath -isnot [string] -or [string]::IsNullOrWhiteSpace($excludedPath)) {
                            throw "$prefix.Exclude.$property must contain nonempty path strings."
                        }
                        $excludedPath = (Get-StringValueHC $excludedPath).Replace('/', '\')
                        if ($excludedPath -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+)' -or $excludedPath -match '[*?]') {
                            throw "$prefix.Exclude.$property requires absolute local or UNC paths without wildcards."
                        }
                        [IO.Path]::GetFullPath($excludedPath).TrimEnd('\')
                    }
                )
            }
        }
        if (($task.Upload.Folders.Count -gt 0 -or $task.PSObject.Properties.Name -contains 'Recurse') -and $task.Recurse -isnot [bool]) {
            throw "$prefix.Recurse must be true or false; it is required when Upload.Folders is nonempty."
        }
        if ($task.SharePoint -isnot [pscustomobject]) { throw "$prefix.SharePoint is required." }
        foreach ($property in $task.SharePoint.PSObject.Properties.Name) {
            if ($property -notin '?', 'SiteUrl', 'DocumentLibraryName', 'FolderPath', 'ClientId', 'TenantId', 'CertificateThumbprint', 'MaxRetries', 'ChunkSizeMB', 'Overwrite') { throw "Unsupported property '$prefix.SharePoint.$property'." }
        }
        if ($task.SharePoint.PSObject.Properties.Name -contains 'Overwrite' -and $task.SharePoint.Overwrite -isnot [bool]) {
            throw "$prefix.SharePoint.Overwrite must be true or false."
        }
        foreach ($property in 'SiteUrl', 'DocumentLibraryName', 'ClientId', 'TenantId', 'CertificateThumbprint') {
            if ($task.SharePoint.$property -isnot [string] -or [string]::IsNullOrWhiteSpace($task.SharePoint.$property)) { throw "$prefix.SharePoint.$property is required." }
        }
        $uri = $null
        if (-not [Uri]::TryCreate($task.SharePoint.SiteUrl, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or $uri.Query -or $uri.Fragment) {
            throw "$prefix.SharePoint.SiteUrl must be an HTTPS site URL without a query or fragment."
        }
        if ($task.SharePoint.FolderPath -match '(^|[\\/])\.{1,2}([\\/]|$)') { throw "$prefix.SharePoint.FolderPath cannot contain '.' or '..' segments." }
        foreach ($property in 'MaxRetries', 'ChunkSizeMB') {
            $value = $task.SharePoint.$property
            $maximum = if ($property -eq 'ChunkSizeMB') { 60 } else { [int]::MaxValue }
            if ($null -ne $value -and ("$value" -notmatch '^\d+$' -or [decimal]$value -lt 1 -or [decimal]$value -gt $maximum)) {
                throw "$prefix.SharePoint.$property must be a whole number between 1 and $maximum."
            }
        }
        $identity = @($task.SharePoint.ClientId, $task.SharePoint.TenantId, $task.SharePoint.CertificateThumbprint) | ConvertTo-Json -Compress
        if ($authentication -and $authentication -ne $identity) { throw 'All tasks must use the same Graph authentication settings. Use separate configurations for different applications.' }
        $authentication = $identity
        $after = $task.AfterUpload
        if ($task.PSObject.Properties.Name -contains 'AfterUpload') {
            if ($after -isnot [pscustomobject]) { throw "$prefix.AfterUpload must be an object." }
            foreach ($property in $after.PSObject.Properties.Name) {
                if ($property -notin '?', 'RemoveFile', 'CopyToFolder') { throw "Unsupported property '$prefix.AfterUpload.$property'." }
            }
            if ($after.PSObject.Properties.Name -contains 'RemoveFile' -and $after.RemoveFile -isnot [bool]) { throw "$prefix.AfterUpload.RemoveFile must be true or false." }
            if ($null -ne $after.CopyToFolder -and $after.CopyToFolder -isnot [string]) { throw "$prefix.AfterUpload.CopyToFolder must be a path string or null." }
        }
        $copyFolder = Get-StringValueHC $after.CopyToFolder
        if ($copyFolder -and $copyFolder -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+)') { throw "$prefix.AfterUpload.CopyToFolder must be an absolute local or UNC path." }
        if ($copyFolder) { $copyFolder = [IO.Path]::GetFullPath($copyFolder).TrimEnd('\') }
        $folderPaths = @()
        $filePaths = @()
        foreach ($selection in $selections) {
            $entry = $selection.Entry
            $sourceProperty = "$prefix.Upload.$($selection.Kind)"
            $uploadType = if ($selection.Kind -eq 'Files') { 'File' } else { 'Folder' }
            if ($entry -is [string]) { $path = $entry; $name = '' }
            elseif ($entry -is [pscustomobject] -and $entry.Path -is [string]) {
                foreach ($property in $entry.PSObject.Properties.Name) {
                    if ($property -notin '?', 'Name', 'Path') { throw "Unsupported property '$sourceProperty.$property'." }
                }
                if ($entry.PSObject.Properties.Name -contains 'Name' -and $entry.Name -isnot [string]) {
                    throw "$sourceProperty entry Name must be a string."
                }
                $path = $entry.Path; $name = $entry.Name
            }
            else { throw "$sourceProperty entries must be strings or objects with Path and optional Name." }
            if ([string]::IsNullOrWhiteSpace($path)) { throw "$sourceProperty requires nonempty paths." }
            $path = (Get-StringValueHC $path).Replace('/', '\')
            if ($path -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+)' -or $path -match '[*?]') {
                throw "$sourceProperty requires absolute local or UNC paths without wildcards."
            }
            $path = [IO.Path]::GetFullPath($path)
            $computer = $task.ComputerName
            if ($computer -and $computer -isnot [string]) { throw "$prefix.ComputerName must be a string or null." }
            $isUnc = $path.StartsWith('\\')
            if (-not $computer -and -not $isUnc) { throw "$prefix.ComputerName is required for local paths; use localhost." }
            if (-not $computer -or $computer -in 'localhost', '.', $env:COMPUTERNAME) { $computer = $env:COMPUTERNAME }
            if ($isUnc -and $computer -ne $env:COMPUTERNAME) { throw "${prefix}: use ComputerName null for UNC paths to avoid a remoting double hop." }
            $target = if ($isUnc) { ($path -split '\\')[2] } else { $computer }
                if ($uploadType -eq 'Folder') { $folderPaths += $path.TrimEnd('\') }
                else { $filePaths += $path }
            $pathPrefix = "$($path.TrimEnd('\'))\"
            $pathExclusions = @{ Attributes = $excludeAttributes }
            foreach ($property in 'Folders', 'Files') {
                $pathExclusions[$property] = @($excludedPaths[$property] | Where-Object {
                    if ($uploadType -eq 'Folder') { $_.StartsWith($pathPrefix, [StringComparison]::OrdinalIgnoreCase) }
                    else { $property -eq 'Files' -and $_ -eq $path }
                    })
            }
            $prepared.Add([pscustomobject]@{
                    TaskIndex = $taskIndex; ComputerName = $computer; Target = $target; Name = $name
                    Path = $path; UploadType = $uploadType; Recurse = [bool]($uploadType -eq 'Folder' -and $task.Recurse); SharePoint = $task.SharePoint
                    Exclude = [pscustomobject]$pathExclusions
                    AfterUpload = [pscustomobject]@{ RemoveFile = [bool]$after.RemoveFile; CopyToFolder = $copyFolder }
                })
        }
        foreach ($property in 'Folders', 'Files') {
            foreach ($excludedPath in $excludedPaths[$property]) {
                $matchingRoots = @($folderPaths | Where-Object {
                        $excludedPath.StartsWith("$_\", [StringComparison]::OrdinalIgnoreCase)
                    })
                if (-not $matchingRoots -and -not ($property -eq 'Files' -and $excludedPath -in $filePaths)) {
                    throw "$prefix.Exclude.$property path '$excludedPath' must be below an Upload.Folders source, or for Exclude.Files match an Upload.Files source in this task."
                }
            }
        }
    }
    foreach ($task in $prepared) {
        foreach ($other in $prepared) {
            if ($task.ComputerName -ne $other.ComputerName) { continue }
            $source = $other.Path.TrimEnd('\')
            $copy = $task.AfterUpload.CopyToFolder
            if ($copy -and ($copy -eq $source -or $copy.StartsWith("$source\", [StringComparison]::OrdinalIgnoreCase) -or
                ($other.UploadType -eq 'File' -and $source.StartsWith("$copy\", [StringComparison]::OrdinalIgnoreCase)))) {
                throw 'AfterUpload.CopyToFolder must be outside every source folder and must not contain a selected source file on the same computer.'
            }
            if (-not [object]::ReferenceEquals($task, $other) -and ($task.Path.TrimEnd('\') -eq $source -or
                ($other.UploadType -eq 'Folder' -and $task.Path.StartsWith("$source\", [StringComparison]::OrdinalIgnoreCase)))) {
                throw 'Source paths must not overlap on the same computer, including duplicate files or files inside selected folders.'
            }
        }
    }
    $prepared.ToArray()
}

function Send-UploadReportHC {
    [CmdletBinding()]
    param ($Settings, [object[]]$Results, [Collections.Generic.List[object]]$SystemErrors, [datetime]$Started, [string]$ConfigurationName)

    $uploadedCount = @($Results | Where-Object Uploaded).Count
    $itemErrorCount = @($Results | Where-Object Error).Count
    $baseLogName = $null
    $attachments = @()
    try {
        $folder = Get-StringValueHC $Settings.SaveLogFiles.Where.Folder
        if ($folder) {
            $logFolder = Get-LogFolderHC $folder
            $baseLogName = Join-Path $logFolder ('{0} - {1} ({2})' -f $Started.ToString('yyyy_MM_dd_HHmmss_fff'), $Settings.ScriptName, $ConfigurationName)
            $days = $Settings.SaveLogFiles.DeleteLogsAfterDays
            if ($days -gt 0) {
                $cutoff = (Get-Date).AddDays(-$days)
                Get-ChildItem -LiteralPath $logFolder -File | Where-Object {
                    $_.Name -match '^\d{4}_\d{2}_\d{2}_\d{6}_\d{3} - ' -and
                    $_.Name -match ' - (Log\.xlsx|Log\.json|System errors log\.json|Mail\.html)$' -and
                    $_.Name.Contains(" - $($Settings.ScriptName) ($ConfigurationName) - ") -and
                    $_.LastWriteTime -lt $cutoff
                } | Remove-Item -Force -ErrorAction Stop
            }
            if ($Results.Count) {
                $excelPath = "$baseLogName - Log.xlsx"
                $overviewColumns = @(
                    'DateTime', 'ComputerName', 'TaskIndex', 'Type', 'Name', 'FileName', 'FullName',
                    'CreationTime', 'LastWriteTime', 'Length', 'Action', 'Uploaded', 'CopiedTo', 'Removed',
                    'Destination', 'WebUrl', 'Error', 'Stage', 'Source'
                )
                $errorColumns = @(
                    'DateTime', 'ComputerName', 'TaskIndex', 'Type', 'Name', 'FullName', 'Action',
                    'Destination', 'Error', 'Stage', 'TargetObject', 'FullyQualifiedErrorId',
                    'ExceptionType', 'ScriptStackTrace', 'PositionMessage'
                )
                $excelParams = @{
                    Path = $excelPath; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true
                    NoNumberConversion = '*'; ErrorAction = 'Stop'
                }
                $Results | Select-Object -Property $overviewColumns |
                    Export-Excel @excelParams -WorksheetName 'Overview' -TableName 'Overview'
                $fileErrors = @($Results | Where-Object Error)
                if ($fileErrors.Count) {
                    $fileErrors | Select-Object -Property $errorColumns |
                        Export-Excel @excelParams -WorksheetName 'Errors' -TableName 'Errors'
                }
                $attachments += $excelPath
            }
        }
    }
    catch { $SystemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Failed saving log files: $_" }) }

    $errorCount = $itemErrorCount + $SystemErrors.Count
    $summary = "$uploadedCount uploaded, $errorCount errors"
    $sendMail = $Settings.SendMail
    $shouldSend = $sendMail.When -eq 'Always' -or ($sendMail.When -eq 'OnError' -and $errorCount) -or ($sendMail.When -eq 'OnErrorOrAction' -and ($errorCount -or $uploadedCount))
    if ($shouldSend) {
        try {
            $title = [System.Net.WebUtility]::HtmlEncode($Settings.ScriptName)
            $rows = $Results | Select-Object ComputerName, Name, Source, Destination, Uploaded, CopiedTo, Removed, Error | ConvertTo-Html -Fragment
            $errorsHtml = ($SystemErrors | ForEach-Object { '<li>{0}</li>' -f [System.Net.WebUtility]::HtmlEncode($_.Message) }) -join ''
            $body = "<html><body><h1>$title</h1><p>$summary</p>$($sendMail.Body)<ul>$errorsHtml</ul>$($rows -join [Environment]::NewLine)</body></html>"
            if ($baseLogName) { Set-Content -LiteralPath "$baseLogName - Mail.html" -Value $body -Encoding utf8 -ErrorAction Stop }
            $mailParams = @{
                From                = Get-StringValueHC $sendMail.From
                SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                SmtpPort            = Get-StringValueHC "$($sendMail.Smtp.Port)"
                MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                Subject             = (@($summary, $sendMail.Subject) | Where-Object { $_ }) -join ', '
                Body                = $body
            }
            foreach ($property in 'To', 'Bcc') { if ($sendMail.$property) { $mailParams[$property] = $sendMail.$property } }
            if ($sendMail.FromDisplayName) { $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName }
            if ($sendMail.Smtp.ConnectionType) { $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType }
            if ($attachments) { $mailParams.Attachments = $attachments }
            if ($errorCount) { $mailParams.Priority = 'High' }
            $username = Get-StringValueHC $sendMail.Smtp.UserName
            $password = Get-StringValueHC $sendMail.Smtp.Password
            if ([bool]$username -ne [bool]$password) { throw 'Both SMTP username and password are required.' }
            if ($username) { $mailParams.Credential = [pscredential]::new($username, (ConvertTo-SecureString $password -AsPlainText -Force)) }
            Send-MailKitMessageHC @mailParams
        }
        catch { $SystemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Failed sending email: $_" }) }
    }
    if ($Settings.SaveInEventLog.Save) {
        try {
            $errorCount = $itemErrorCount + $SystemErrors.Count
            $events = @([pscustomobject]@{ Message = 'Script started'; DateTime = $Started; EntryType = 'Information'; EventID = 100 })
            $events += [pscustomobject]@{ Message = "$uploadedCount uploaded, $errorCount errors"; DateTime = Get-Date; EntryType = $(if ($errorCount) { 'Error' } else { 'Information' }); EventID = 101 }
            $events += @($SystemErrors | ForEach-Object { [pscustomobject]@{ Message = $_.Message; DateTime = $_.DateTime; EntryType = 'Error'; EventID = 102 } })
            Write-EventsToEventLogHC -Source $Settings.ScriptName -LogName (Get-StringValueHC $Settings.SaveInEventLog.LogName) -Events $events
        }
        catch { $SystemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Failed writing event log: $_" }) }
    }
    if ($baseLogName -and $SystemErrors.Count) {
        try {
            ConvertTo-Json -InputObject $SystemErrors.ToArray() -Depth 5 |
            Set-Content -LiteralPath "$baseLogName - System errors log.json" -Encoding utf8 -ErrorAction Stop
        }
        catch { $SystemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Failed saving JSON log: $_" }) }
    }
}