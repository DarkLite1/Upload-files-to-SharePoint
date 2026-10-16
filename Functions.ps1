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
            if ($property -notin '?', 'SiteUrl', 'DocumentLibraryName', 'ClientId', 'TenantId', 'CertificateThumbprint', 'MaxRetries', 'ChunkSizeMB', 'Overwrite') { throw "Unsupported property '$prefix.SharePoint.$property'." }
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
            $sharePointFileName = $null
            if ($entry -is [pscustomobject]) {
                foreach ($property in $entry.PSObject.Properties.Name) {
                    if ($property -notin '?', 'Name', 'Source', 'Destination' -and -not ($uploadType -eq 'File' -and $property -eq 'SharePointFileName')) {
                        throw "Unsupported property '$sourceProperty.$property'."
                    }
                }
                if ($entry.PSObject.Properties.Name -contains 'Name' -and $entry.Name -isnot [string]) {
                    throw "$sourceProperty entry Name must be a string."
                }
                if ($entry.Source -isnot [string] -or [string]::IsNullOrWhiteSpace($entry.Source)) {
                    throw "$sourceProperty.Source must be a nonempty path string."
                }
                if ($entry.PSObject.Properties.Name -notcontains 'Destination') {
                    throw "$sourceProperty.Destination is required; use null or an empty string for the library root."
                }
                if ($null -ne $entry.Destination -and $entry.Destination -isnot [string]) {
                    throw "$sourceProperty.Destination must be a string or null; use null or an empty string for the library root."
                }
                if ($entry.Destination.Length -and [string]::IsNullOrWhiteSpace($entry.Destination)) {
                    throw "$sourceProperty.Destination cannot contain only whitespace; use an empty string for the library root."
                }
                if ($entry.Destination -match '(^|[\\/])\.{1,2}([\\/]|$)') {
                    throw "$sourceProperty.Destination cannot contain '.' or '..' segments."
                }
                if ($entry.PSObject.Properties.Name -contains 'SharePointFileName') {
                    $sharePointFileName = $entry.SharePointFileName
                    if ($null -eq $sharePointFileName -or ($sharePointFileName -is [string] -and [string]::IsNullOrWhiteSpace($sharePointFileName))) {
                        $sharePointFileName = $null
                    }
                    elseif ($sharePointFileName -isnot [string] -or
                        $sharePointFileName.Length -gt 255 -or $sharePointFileName -match '["*:<>?/\\|\x00-\x1f]' -or
                        $sharePointFileName.StartsWith('.') -or $sharePointFileName.EndsWith('.') -or $sharePointFileName.Trim() -ne $sharePointFileName) {
                        throw "$sourceProperty.SharePointFileName must be null, blank, or a filename of at most 255 characters, without paths, invalid SharePoint characters, or leading/trailing periods or whitespace."
                    }
                }
                $path = $entry.Source; $name = $entry.Name
            }
            else { throw "$sourceProperty entries must be objects with Source and Destination, and optional Name." }
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
                    Path = $path; Destination = $entry.Destination; UploadType = $uploadType; Recurse = [bool]($uploadType -eq 'Folder' -and $task.Recurse); SharePoint = $task.SharePoint
                    SharePointFileName = $sharePointFileName
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

function Get-MailThemeHC {
    @{
        StatusError = '#fee2e2'; AccentError = '#dc2626'; AccentSuccess = '#16a34a'
        AccentIdle = '#6b7280'; AccentSystem = '#7c2d12'
        GradError = @('#7f1d1d', '#dc2626'); GradSuccess = @('#14532d', '#16a34a')
        GradIdle = @('#374151', '#6b7280')
        TextMain = '#111827'; TextMuted = '#374151'; TextLight = '#6b7280'
        BgPage = '#e5e7eb'; BgWhite = '#ffffff'; BorderMain = '#d1d5db'; BorderLight = '#e5e7eb'
        LinkColor = '#2563eb'; BodyWidth = 620
        FontStack = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif"
        MonoStack = 'Consolas, Menlo, monospace'
    }
}

function New-PillHtmlHC {
    param ([string]$Text, [string]$Bg, [string]$Color = '#ffffff')

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $encoded = [System.Net.WebUtility]::HtmlEncode($Text)
    $upper = [System.Net.WebUtility]::HtmlEncode($Text.ToUpperInvariant())
    $width = [int][Math]::Ceiling(($Text.Length * 8.5) + 26)
    $span = "<span style=`"display:inline-block; padding:3px 10px; background-color:$Bg; color:$Color; border-radius:12px; font-size:11px; font-weight:700; letter-spacing:0; text-transform:uppercase; line-height:1.6;`">$encoded</span>"
    $vml = "<v:roundrect xmlns:v=`"urn:schemas-microsoft-com:vml`" xmlns:w=`"urn:schemas-microsoft-com:office:word`" arcsize=`"50%`" fillcolor=`"$Bg`" stroked=`"f`" style=`"height:26px; width:${width}px; v-text-anchor:middle; mso-padding-alt:0;`">" +
        '<w:anchorlock/>' +
        "<center style=`"color:$Color; font-family:sans-serif; font-size:11px; font-weight:700; letter-spacing:0;`">$upper</center></v:roundrect>"
    "<!--[if mso]>$vml<![endif]--><!--[if !mso]><!-->$span<!--<![endif]-->"
}

function Build-MailSummaryBannerHC {
    param ([int]$Uploaded, [int]$Errors)

    $theme = Get-MailThemeHC
    $color = if ($Uploaded) { $theme.AccentSuccess } else { $theme.AccentIdle }
    $pills = "<td valign='middle' style='vertical-align:middle; padding:4px 6px 4px 0; white-space:nowrap; font-size:0;'>$(New-PillHtmlHC -Text "$Uploaded Uploaded" -Bg $color)</td>"
    if ($Errors) {
        $label = "$Errors Error" + $(if ($Errors -ne 1) { 's' })
        $pills += "<td valign='middle' style='vertical-align:middle; padding:4px 6px 4px 0; white-space:nowrap; font-size:0;'>$(New-PillHtmlHC -Text $label -Bg $theme.AccentError)</td>"
    }
    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 16px 0;">
    <tr><td style="padding:0;">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
            <tr><td valign="middle" style="vertical-align:middle; padding:4px 12px 4px 0; font-size:13px; font-weight:600; color:$($theme.TextMain);">Summary</td>$pills</tr>
        </table>
    </td></tr>
</table>
<!--[if mso]><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table><![endif]-->
"@
}

function ConvertTo-FileUrlHC {
    param ([string]$Path)

    $uri = $null
    if ([Uri]::TryCreate($Path, [UriKind]::Absolute, [ref]$uri) -and $uri.IsFile) { $uri.AbsoluteUri }
    else { '' }
}

function Get-TaskDescriptionHC {
    param ([pscustomobject]$Task)

    $description = if ($Task.UploadType -eq 'File') { 'Upload file' }
    elseif ($Task.UploadType -eq 'Folder') { 'Upload folder contents' }
    else { 'Upload files' }
    if ($Task.UploadType -eq 'Folder' -and $Task.Recurse) { $description += ', including subfolders' }
    if ($Task.SharePoint.SiteUrl) {
        $destination = (@($Task.SharePoint.SiteUrl.TrimEnd('/'), $Task.SharePoint.DocumentLibraryName, $Task.Destination) | Where-Object { $_ }) -join '/'
        $description += " to $destination"
    }
    if ($Task.SharePointFileName) { $description += " as '$($Task.SharePointFileName)'" }
    if ($Task.AfterUpload.CopyToFolder) { $description += ", copy to $($Task.AfterUpload.CopyToFolder)" }
    if ($Task.AfterUpload.RemoveFile) { $description += ', then remove the source' }
    $excluded = @()
    foreach ($property in 'Folders', 'Files') {
        $count = @($Task.Exclude.$property | Where-Object { $_ }).Count
        if ($count) { $excluded += "$count $($property.ToLowerInvariant())" }
    }
    if ($Task.Exclude.Attributes) { $excluded += ($Task.Exclude.Attributes -join '/') + ' items' }
    if ($excluded) { $description += ', excluding ' + ($excluded -join ' and ') }
    $description
}

function Get-UploadMailJobsHC {
    param ([object[]]$Tasks = @(), [object[]]$Results = @())

    $groups = if ($Tasks.Count) {
        foreach ($group in ($Tasks | Group-Object TaskIndex)) {
            $taskIndex = $group.Group[0].TaskIndex
            [pscustomobject]@{ Sources = $group.Group; Results = @($Results | Where-Object { $_.TaskIndex -eq $taskIndex }) }
        }
    }
    else {
        foreach ($group in ($Results | Group-Object ComputerName, TaskIndex, Name)) {
            $sources = $group.Group | Select-Object ComputerName, Name, @{ Name = 'Path'; Expression = { $_.Source } }, Destination, WebUrl -Unique
            [pscustomobject]@{ Sources = @($sources); Results = $group.Group }
        }
    }
    foreach ($group in $groups) {
        $entries = @(
            foreach ($source in $group.Sources) {
                $computer = $source.ComputerName
                $linkPath = $source.Path
                if ($source.Path -match '^\\\\([^\\]+)') { $computer = $Matches[1] }
                elseif ($source.Path -match '^[A-Za-z]:\\' -and $computer) {
                    $linkComputer = if ($computer -in 'localhost', '.') { $env:COMPUTERNAME } else { $computer }
                    $linkPath = $source.Path -replace '^(.):', ('\\{0}\$1$' -f $linkComputer)
                }
                [pscustomobject]@{
                    ComputerName = $computer; Name = $source.Name; Path = $source.Path; LinkPath = $linkPath
                    Destination = $source.Destination; WebUrl = $source.WebUrl
                }
            }
        )
        [pscustomobject]@{
            ComputerName = ($entries.ComputerName | Sort-Object -Unique) -join ', '
            Entries = $entries; Path = $entries[0].Path
            Description = (($group.Sources | ForEach-Object { Get-TaskDescriptionHC $_ }) | Select-Object -Unique) -join '; '
            Uploaded = @($group.Results | Where-Object Uploaded).Count
            Copied = @($group.Results | Where-Object CopiedTo).Count
            Removed = @($group.Results | Where-Object Removed).Count
            Errors = @($group.Results | Where-Object Error).Count
        }
    }
}

function Build-MailSystemErrorsBlockHC {
    param ([string[]]$Message)

    $messages = @($Message | Where-Object { $_ })
    if (-not $messages.Count) { return '' }
    $theme = Get-MailThemeHC
    @"
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; margin:0 0 20px 0; table-layout:fixed;">
    <tr><td style="padding:0 0 8px 0; font-size:11px; font-weight:700; color:$($theme.TextLight); text-transform:uppercase;">System Errors ($($messages.Count))</td></tr>
</table>
<!--[if mso]><table role="presentation" width="100%"><tr><td height="12" style="font-size:0; line-height:12px; mso-line-height-rule:exactly;">&#160;</td></tr></table><![endif]-->
"@
}

function Build-MailJobRowHC {
    param ([pscustomobject]$Job)

    $theme = Get-MailThemeHC
    $accent = if ($Job.Errors) { $theme.AccentError } elseif ($Job.Uploaded) { $theme.AccentSuccess } else { $theme.AccentIdle }
    $titleHtml = @(foreach ($entry in $Job.Entries) {
        $href = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $entry.LinkPath))
        $path = [System.Net.WebUtility]::HtmlEncode($entry.Path)
        $title = if ($href) { "<a href='$href' target='_blank' rel='noopener noreferrer' style='color:$($theme.TextMain); text-decoration:none;'>$path</a>" } else { $path }
        "<div style='margin:0; font-weight:700; color:$($theme.TextMain); font-size:13px; line-height:16px; mso-line-height-rule:exactly; overflow-wrap:anywhere; word-break:break-all;'>$title</div>"
    }) -join ''
    $resultText = '{0}&nbsp;uploaded' -f $Job.Uploaded
    if ($Job.Errors) { $resultText += '<br>{0}&nbsp;error{1}' -f $Job.Errors, $(if ($Job.Errors -ne 1) { 's' }) }
    @"
<table class="task-row" role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:separate; margin:0 0 4px 0; table-layout:fixed; background-color:$($theme.BgWhite); border:1px solid $($theme.BorderMain); border-left:3px solid $accent; border-radius:6px;">
    <tr><td class="task-content" valign="middle" style="padding:8px 12px;">$titleHtml</td>
    <td class="task-result" valign="middle" align="right" width="90" style="padding:8px 12px; color:$($theme.TextMuted); font-size:11px; line-height:15px; mso-line-height-rule:exactly; white-space:nowrap;">$resultText</td></tr>
</table>
"@
}

function Build-MailComputerCardHC {
    param ([string]$ComputerName, [pscustomobject[]]$Job)

    $theme = Get-MailThemeHC
    $uploaded = ($Job | Measure-Object -Property Uploaded -Sum).Sum
    $errors = ($Job | Measure-Object -Property Errors -Sum).Sum
    $symbol, $gradient = if ($errors) { '&#10006;', $theme.GradError }
    elseif ($uploaded) { '&#10003;', $theme.GradSuccess }
    else { '&#8211;', $theme.GradIdle }
    $gradFrom, $gradTo = $gradient
    $gradMid = '#' + ((0, 2, 4 | ForEach-Object {
        '{0:x2}' -f [int][Math]::Round((
            [Convert]::ToInt32($gradFrom.Substring(1 + $_, 2), 16) +
            [Convert]::ToInt32($gradTo.Substring(1 + $_, 2), 16)
        ) / 2)
    }) -join '')
    $headerLabel = '{0}&nbsp;uploaded' -f $uploaded
    if ($errors) { $headerLabel += ' &middot; {0}&nbsp;error{1}' -f $errors, $(if ($errors -ne 1) { 's' }) }
    $rows = ($Job | Sort-Object @{ Expression = { if ($_.Errors) { 0 } elseif ($_.Uploaded) { 1 } else { 2 } } }, Path, Description |
        ForEach-Object { Build-MailJobRowHC $_ }) -join '<!--[if mso]><table role="presentation" width="100%" bgcolor="#ffffff"><tr><td height="4" style="font-size:0; line-height:4px; mso-line-height-rule:exactly;">&#160;</td></tr></table><![endif]-->'
    $pathCount = @($Job.Entries.Path | Sort-Object -Unique).Count
    @"
<table class="computer-card" role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($theme.BgWhite)" style="border-collapse:separate; margin:0 0 16px 0; table-layout:fixed; background-color:$($theme.BgWhite); border:1px solid $($theme.BorderLight); border-radius:10px; overflow:hidden; box-shadow:0 2px 4px rgba(0,0,0,0.06);">
    <tr><td bgcolor="$gradMid" style="padding:0; background-color:$gradMid; background-image:linear-gradient(135deg, $gradFrom 0%, $gradTo 100%);">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse; table-layout:fixed;">
            <tr><td valign="middle" align="center" width="52" style="padding:12px 0; font-size:18px; font-weight:bold; color:#ffffff; line-height:24px;">$symbol</td>
            <td valign="middle" style="padding:12px 8px 12px 0; overflow-wrap:anywhere; word-break:break-word;"><p style="margin:0; font-size:16px; font-weight:700; color:#ffffff; line-height:20px;">$([System.Net.WebUtility]::HtmlEncode($ComputerName))</p><p style="margin:2px 0 0 0; font-size:12px; color:#f1f2f4; line-height:17px;">$pathCount path$(if ($pathCount -ne 1) { 's' })</p></td>
            <td class="computer-total" valign="middle" align="right" width="140" style="padding:12px 14px 12px 6px; font-size:12px; font-weight:700; color:#e5e7eb; text-transform:uppercase;">$headerLabel</td></tr>
        </table>
    </td></tr>
    <tr><td class="card-content" style="padding:12px 16px;">$rows</td></tr>
</table>
<!--[if mso]><table role="presentation" width="100%"><tr><td height="16" style="font-size:0; line-height:16px; mso-line-height-rule:exactly;">&#160;</td></tr></table><![endif]-->
"@
}

function Get-MailBodyHtmlHC {
    param (
        [string]$ScriptName, [string]$Body, [pscustomobject[]]$Job = @(),
        [int]$Uploaded, [int]$Errors, [string[]]$SystemError,
        [string]$LogFolderPath, [bool]$HasAttachments,
        [datetime]$ScriptStartTime, [datetime]$ScriptEndTime = (Get-Date), [string]$BrowserViewFilePath
    )

    $theme = Get-MailThemeHC
    $banner = Build-MailSummaryBannerHC -Uploaded $Uploaded -Errors $Errors
    $systemErrorsBlock = Build-MailSystemErrorsBlockHC $SystemError
    $cards = $Job | Group-Object ComputerName | Sort-Object @{
        Expression = {
            if (($_.Group | Measure-Object Errors -Sum).Sum) { 0 }
            elseif (($_.Group | Measure-Object Uploaded -Sum).Sum) { 1 }
            else { 2 }
        }
    }, Name | ForEach-Object { Build-MailComputerCardHC -ComputerName $_.Name -Job $_.Group }
    $linkStyle = "color:$($theme.LinkColor); text-decoration:none; font-weight:600;"
    $browserViewRow = ''
    if ($BrowserViewFilePath) {
        $browserUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $BrowserViewFilePath))
        $browserTitle = [System.Net.WebUtility]::HtmlEncode($BrowserViewFilePath)
        $browserViewRow = "<!--[if mso]><tr><td style='padding:0 0 8px 0; font-size:12px; color:$($theme.TextMuted);'>If this mail is not visible, please <a href='$browserUrl' title='$browserTitle' style='$linkStyle'>click here to view it in the browser</a>.</td></tr><![endif]-->"
    }
    $links = @()
    if ($LogFolderPath) {
        $folderPath = [System.Net.WebUtility]::HtmlEncode($LogFolderPath)
        $folderUrl = [System.Net.WebUtility]::HtmlEncode((ConvertTo-FileUrlHC $LogFolderPath))
        $links += "<!--[if mso]><a href='$folderPath' title='$folderPath' style='$linkStyle'>Open log folder</a><![endif]--><!--[if !mso]><!--><a href='$folderUrl' title='$folderPath' style='$linkStyle'>Open log folder</a><!--<![endif]-->"
    }
    if ($HasAttachments) { $links += 'details in the attachments' }
    $linksBlock = if ($links) { "<p style='margin:0 0 14px 0; font-size:12px; color:$($theme.TextMuted);'>$($links -join ' &middot; ')</p>" }
    $duration = $ScriptEndTime - $ScriptStartTime
    $durationText = '{0:00}:{1:00}:{2:00}' -f [int][Math]::Floor($duration.TotalHours), $duration.Minutes, $duration.Seconds
    $footer = @"
<table class="timing" role="presentation" align="center" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin:16px auto 0 auto; color:$($theme.TextLight); font-size:11px;">
    <tr><td style="padding:0 20px 0 0;"><strong>Started</strong> $($ScriptStartTime.ToString('dd/MM/yyyy HH:mm'))</td><td style="padding:0 20px 0 0;"><strong>Ended</strong> $($ScriptEndTime.ToString('dd/MM/yyyy HH:mm'))</td><td><strong>Duration</strong> $durationText</td></tr>
</table>
<p style="margin:6px 0 0 0; text-align:center; color:$($theme.TextLight); font-size:11px; font-family:$($theme.MonoStack); overflow-wrap:anywhere;">$([System.Net.WebUtility]::HtmlEncode("$env:USERDNSDOMAIN\$env:USERNAME on $env:COMPUTERNAME, PowerShell $($PSVersionTable.PSVersion)"))</p>
"@
    $noJobs = if (-not $cards) { "<p style='margin:0; color:$($theme.TextLight); font-style:italic;'>No tasks were executed.</p>" }
    @"
<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<!--[if mso]><style type="text/css">v\:* { behavior:url(#default#VML); display:inline-block; } table { mso-table-lspace:0pt; mso-table-rspace:0pt; } td { mso-line-height-rule:exactly; }</style><![endif]-->
<style type="text/css">
body { font-family:$($theme.FontStack); font-size:13px; color:$($theme.TextMain); background-color:$($theme.BgPage); margin:0; padding:0; }
a { color:$($theme.LinkColor); text-decoration:none; }
h1 { font-size:22px; font-weight:700; color:$($theme.TextMain); margin:0 0 4px 0; overflow-wrap:anywhere; }
table { border-collapse:collapse; }
</style>
<!--[if !mso]><!--><style type="text/css">
table.mail-root { max-width:900px !important; }
@media only screen and (max-width:600px) {
    td.page-padding { padding:12px !important; }
    td.card-content { padding:8px !important; }
    td.status-cell { display:none !important; }
    td.computer-total { width:100px !important; }
    td.task-result { width:75px !important; }
    table.timing td { display:block !important; padding:2px 0 !important; text-align:center; }
}
</style><!--<![endif]-->
</head>
<body style="margin:0; padding:0;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="$($theme.BgPage)" style="background-color:$($theme.BgPage);">
    <tr><td class="page-padding" align="center" valign="top" bgcolor="$($theme.BgPage)" style="padding:20px;">
        <!--[if mso]><table role="presentation" cellpadding="0" cellspacing="0" border="0" width="$($theme.BodyWidth)" align="center"><tr><td><![endif]-->
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" class="mail-root" style="width:100%; margin:0 auto; text-align:left; table-layout:fixed;">
            <tr><td style="padding:0 0 4px 0;"><h1>$([System.Net.WebUtility]::HtmlEncode($ScriptName))</h1></td></tr>
            <tr><td style="padding:0 0 16px 0; color:$($theme.TextMuted); font-size:13px; line-height:1.6; overflow-wrap:anywhere;">$Body</td></tr>
            $browserViewRow
            <tr><td>$linksBlock</td></tr>
            <tr><td>$banner</td></tr>
            <tr><td>$systemErrorsBlock</td></tr>
            <tr><td>$($cards -join '')$noJobs</td></tr>
            <tr><td>$footer</td></tr>
        </table>
        <!--[if mso]></td></tr></table><![endif]-->
    </td></tr>
</table>
</body></html>
"@
}

function Send-UploadReportHC {
    [CmdletBinding()]
    param ($Settings, [object[]]$Results, [Collections.Generic.List[object]]$SystemErrors, [datetime]$Started, [string]$ConfigurationName, [object[]]$Tasks = @())

    $uploadedCount = @($Results | Where-Object Uploaded).Count
    $itemErrorCount = @($Results | Where-Object Error).Count
    $baseLogName = $null
    $logFolder = $null
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
            $mailHtmlPath = if ($baseLogName) { "$baseLogName - Mail.html" }
            $bodyParams = @{
                ScriptName = $Settings.ScriptName; Body = $sendMail.Body
                Job = @(Get-UploadMailJobsHC -Tasks $Tasks -Results $Results)
                Uploaded = $uploadedCount; Errors = $errorCount; SystemError = @($SystemErrors.Message)
                LogFolderPath = $logFolder; HasAttachments = [bool]$attachments
                ScriptStartTime = $Started; ScriptEndTime = Get-Date; BrowserViewFilePath = $mailHtmlPath
            }
            $body = Get-MailBodyHtmlHC @bodyParams
            if ($mailHtmlPath) {
                try { Set-Content -LiteralPath $mailHtmlPath -Value $body -Encoding utf8 -NoNewline -ErrorAction Stop }
                catch {
                    $SystemErrors.Add([pscustomobject]@{ DateTime = Get-Date; Message = "Failed saving email HTML: $_" })
                    $errorCount = $itemErrorCount + $SystemErrors.Count
                    $summary = "$uploadedCount uploaded, $errorCount errors"
                    $bodyParams.Errors = $errorCount
                    $bodyParams.SystemError = @($SystemErrors.Message)
                    $bodyParams.BrowserViewFilePath = $null
                    $body = Get-MailBodyHtmlHC @bodyParams
                }
            }
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