# Upload Files to SharePoint

Upload files from a Windows folder or network share to a SharePoint document library. After a successful upload, you can keep the source file, copy it to an archive, remove it, or archive it and then remove it.

## 1. Install Requirements

Use **PowerShell 7**, not Windows PowerShell 5.1. If needed, install it from a terminal:

```powershell
winget install --id Microsoft.PowerShell --source winget
```

Open PowerShell 7 (`pwsh`) and check the version:

```powershell
$PSVersionTable.PSVersion
```

Install the required modules under the Windows account that will run the uploads:

```powershell
Install-Module -Name Microsoft.Graph.Authentication -Repository PSGallery -Scope CurrentUser
Install-Module -Name ImportExcel -Repository PSGallery -Scope CurrentUser
```

Accept the NuGet provider or PSGallery trust prompt if shown. These commands require access to PowerShell Gallery. `Microsoft.Graph.Authentication` connects to SharePoint; `ImportExcel` creates Excel reports and is required even when logging is disabled.

Verify both modules load:

```powershell
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
```

For installation for all accounts, open PowerShell 7 as administrator and use `-Scope AllUsers` instead. For a scheduled task, make sure its account can access the modules and launch it with `pwsh.exe`.

## 2. Set Up SharePoint Access

Ask your Microsoft 365 administrator for:

- An Entra ID application with a **client ID** and **tenant ID**.
- Certificate-based authentication, with the certificate and its private key installed and accessible to the Windows account running the uploads. You will need its **certificate thumbprint**.
- Microsoft Graph **application** permission `Sites.ReadWrite.All`, or `Sites.Selected` with a separate **write** grant on the destination site. Application permissions require administrator consent.
- The SharePoint site URL and document library display name, such as `https://contoso.sharepoint.com/sites/IT` and `Documents`.

The Windows account also needs permission to read the source files, write to the archive when copying, and delete source files when removal is enabled.

## 3. Configure Your Upload

Keep all four PowerShell files together. Use [Example.json](Example.json) as the configuration template; its `?` section explains the available fields.

For your first run, keep only the first example task and change:

| Setting | What to enter |
| --- | --- |
| `Tasks[].ComputerName` | `localhost` for a local folder, or `null` for a network share. |
| `Tasks[].Folders` | Your source folder paths. Entries can be strings or objects with `Name` and `Path`. |
| `Tasks[].Recurse` | `true` to include subfolders, preserving their relative paths; otherwise `false`. |
| `Tasks[].Exclude.Attributes` | Optional array: `["Hidden"]`, `["System"]`, or `["Hidden", "System"]` to skip those files and folder trees. Use `[]` or omit the setting to include them. |
| `Tasks[].SharePoint` | Your site URL, library name, optional destination `FolderPath`, client ID, tenant ID, and certificate thumbprint. |
| `Tasks[].SharePoint.Overwrite` | `false` (default) reports an error if the destination file exists. Set to `true` only when you want to replace existing SharePoint files. |
| `Settings.ScriptName` | A name for your reports and log files. |
| `Settings.SaveLogFiles.Where.Folder` | A dedicated log folder outside your source folders, or `null` to disable log files. |

Replace the `ENV:...` authentication values with your actual IDs and thumbprint, or define the named environment variables for the account doing the uploads. Use identical authentication settings in all tasks.

Start with `Settings.SendMail.When` set to `Never`, `Settings.SaveInEventLog.Save` set to `false`, and these post-upload settings:

```json
"AfterUpload": {
  "RemoveFile": false,
  "CopyToFolder": null
}
```

**The second task in the example removes source files. Remove that task from your configuration unless you intend to use it.** There is no preview/WhatIf mode.

## 4. Run the Upload

From the folder containing the PowerShell files, run:

```powershell
& '.\Main.ps1' -ConfigurationJsonFile '.\Example.json'
```

Check the uploaded files in SharePoint and review the returned results before enabling removal. Errors are shown in each result's `Error` field. Configuration or reporting errors also produce warnings. A failed run exits with code `1`.

## Copy or Remove After Upload

Set `AfterUpload` separately for each task:

| Desired result | `RemoveFile` | `CopyToFolder` |
| --- | --- | --- |
| Keep the source | `false` | `null` |
| Remove the source | `true` | `null` |
| Copy to an archive and keep the source | `false` | Absolute archive folder path |
| Archive, then remove the source | `true` | Absolute archive folder path |

For example, to archive successful uploads and then remove their source files:

```json
"AfterUpload": {
  "RemoveFile": true,
  "CopyToFolder": "C:\\Archive\\Reports"
}
```

Failed uploads never trigger copying or removal. If copying fails, the source stays in place. Existing archive files are not overwritten; a name collision is reported as an error. Missing archive folders are created. Archive folders must be outside all source folders, and source folders must not overlap.

## Exclude Hidden or System Files

Add this to a task to skip both attributes:

```json
"Exclude": {
  "Attributes": ["Hidden", "System"]
}
```

A file matching either attribute is not uploaded, copied, or removed. A matching folder, including a source folder, is skipped together with everything inside it. Excluded items are not errors and do not appear as uploads in the reports. Access errors on non-excluded items are still reported.

Use only `"Hidden"` or only `"System"` to exclude one attribute. Omit `Exclude`, use `{}`, or set `Attributes` to `[]` to include both. The first example task excludes both attributes.

## Optional Reporting

- **Log files:** `Settings.SaveLogFiles` controls the log folder and retention. Excel records every file upload and file/task error, including conflicts and failed copying or removal. A separate `System errors log.json` contains only system errors, such as configuration or reporting failures, and is created only when those errors occur. File results are not duplicated in JSON. `DeleteLogsAfterDays: 0` keeps logs indefinitely.
- **Email:** `Settings.SendMail.When` supports `Never`, `Always`, `OnError`, and `OnErrorOrAction`. To enable it, configure the sender, recipients, SMTP server, and `AssemblyPath` values. MailKit and MimeKit are .NET libraries, not PowerShell modules; compatible DLLs and their dependencies must be installed separately. Leave email disabled until these are available. The email HTML is also saved when file logging is enabled.
- **Windows event log:** Set `Settings.SaveInEventLog.Save` to `true` and specify `LogName`. Ask an administrator to register the event source named by `Settings.ScriptName` if the execution account cannot create it.

## Remote Folders

For a folder on another computer, set `ComputerName` to that computer's name and use a path local to it. That computer needs PowerShell 7, `Microsoft.Graph.Authentication`, and access to the certificate/private key. Install the module there using the command above under the executing account, or install it for all users.

PowerShell remoting must be configured and accessible. `PSSessionConfiguration` defaults to `PowerShell.7`. Prefer an archive folder local to the remote computer; accessing a network archive from a remote session can require additional authentication setup.

For network shares, use `ComputerName: null` and UNC paths such as `\\server\share\Reports`. Files are then uploaded by the computer running the command, without remoting.

## Important Behavior

- Existing SharePoint files cause an error unless `SharePoint.Overwrite` is `true`. On a conflict, the source is neither copied nor removed. Missing destination folders are created.
- Source files you keep are attempted again on the next run. Existing destinations then cause errors unless you enable overwriting. Avoid matching destination filenames across tasks unless overwriting is intended.
- Files must be finished writing before upload. Do not run overlapping instances against the same source folders.
- Hidden and system files are included unless excluded with `Exclude.Attributes`. Folder links are not followed, and linked source folders or files are rejected. Empty folders are not uploaded; local folders are never removed.
- `MaxConcurrent.JobsTotal` limits simultaneous folder jobs; `JobsPerComputer` limits them per computer or source file server. Files within each folder are processed one at a time.