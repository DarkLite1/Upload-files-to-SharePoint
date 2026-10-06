# Upload Files to SharePoint

Upload selected files or the contents of Windows folders and network shares to a SharePoint document library. After a successful upload, you can keep the source file, copy it to an archive, remove it, or archive it and then remove it.

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
| `Tasks[].ComputerName` | `localhost` for local files/folders, or `null` for a network share. |
| `Tasks[].Upload.Folders` | Folders whose contents should be uploaded. Entries can be path strings or objects with `Name` and `Path`. |
| `Tasks[].Upload.Files` | Specific files to upload. Object entries can add `SharePointFileName` to use a different name on SharePoint. Other files in their parent folders are not uploaded. |
| `Tasks[].Recurse` | Required for folder selections: `true` includes subfolders and preserves their relative paths; `false` selects only direct contents. Optional and unused for file-only tasks. |
| `Tasks[].Exclude.Attributes` | Optional array: `["Hidden"]`, `["System"]`, or `["Hidden", "System"]` to skip those files and folder trees. Use `[]` or omit the setting to include them. |
| `Tasks[].Exclude.Folders` | Optional array of absolute subfolder paths to skip with all their contents. |
| `Tasks[].Exclude.Files` | Optional array of absolute file paths to skip. |
| `Tasks[].SharePoint` | Your site URL, library name, optional destination `FolderPath`, client ID, tenant ID, and certificate thumbprint. |
| `Tasks[].SharePoint.Overwrite` | `false` (default) reports an error if the destination file exists. Set to `true` only when you want to replace existing SharePoint files. |
| `Settings.ScriptName` | A name for your reports and log files. |
| `Settings.SaveLogFiles.Where.Folder` | A dedicated log folder outside your source folders, or `null` to disable log files. |

Replace the `ENV:...` authentication values with your actual IDs and thumbprint, or define the named environment variables for the account doing the uploads. Use identical authentication settings in all tasks.

Use `Upload.Folders`, `Upload.Files`, or both in each task. For example:

```json
"Upload": {
  "Folders": ["C:\\Exports\\Reports"],
  "Files": [{ "Name": "Manifest", "Path": "C:\\Exports\\manifest.csv" }]
}
```

Either list may be omitted or set to `[]`, but at least one source is required. Paths must be absolute local or UNC paths, without wildcards; `ENV:` paths resolve on the script host. `Name` is an optional report label, not a destination filename. The former top-level `Tasks[].Folders` setting is no longer accepted.

Folder contents go into `SharePoint.FolderPath` without the selected root folder's name; `Recurse: true` preserves subfolders. Explicit files go directly into `SharePoint.FolderPath` by filename. Do not select the same file twice, overlapping folders, or explicit files inside selected folders, even when `Recurse` is false.

To give a selected file a different name on SharePoint, use an object entry:

```json
"Upload": {
  "Files": [
    {
      "Path": "C:\\Exports\\manifest.csv",
      "SharePointFileName": "Latest manifest.csv"
    }
  ]
}
```

`SharePointFileName` is optional and supported only in `Upload.Files`. It is the complete destination filename, including the desired extension, not a folder path. Omit it or set it to `null`, `""`, or whitespace only to keep the original filename. A nonblank value must be a literal filename of at most 255 characters, without invalid SharePoint characters or leading/trailing periods or whitespace. `Name` remains a separate report label. The local file and archive copy keep their original names; Excel's `Destination` shows the SharePoint name. `SharePoint.Overwrite` applies to that destination name. This uploads directly under the chosen name; it does not rename an existing SharePoint item.

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

Failed uploads never trigger copying or removal. If copying fails, the source stays in place. Existing archive files are not overwritten; a name collision is reported as an error. Missing archive folders are created. Archive folders must be outside all selected source folders and must not contain any explicitly selected source file on the same computer. Folder selections preserve relative paths in the archive; explicit files are copied directly into the archive by filename.

## Exclusions

Use separate lists for attributes, folder paths, and file paths. For a source folder `C:\Exports\Reports`, for example:

```json
"Exclude": {
  "Attributes": ["Hidden", "System"],
  "Folders": ["C:\\Exports\\Reports\\Drafts"],
  "Files": ["C:\\Exports\\Reports\\status.json"]
}
```

`Attributes` accepts `Hidden`, `System`, or both. Items matching either selected attribute are skipped. For folder selections, an attribute-excluded folder, including a source root, protects its entire contents. Explicit file selections check the selected file's own attributes.

`Exclude.Folders` excludes specific subfolder trees below an `Upload.Folders` source in the same task. `Exclude.Files` excludes specific files below an `Upload.Folders` source or matching an explicit `Upload.Files` selection. Use absolute local or UNC paths. Matching is case-insensitive and literal, not wildcard-based. `ENV:` values resolve on the computer running the command.

Excluded files are never uploaded, copied, or removed. Excluded folders are not entered. Excluded items are not errors and do not appear as uploads in the reports. Access errors on non-excluded items are still reported.

All three lists are optional. Omit a list or set it to `[]` when unused; omit `Exclude` or use `{}` for no exclusions. The first example task excludes Hidden and System attributes and leaves the file and folder lists empty.

## Optional Reporting

- **Log files:** `Settings.SaveLogFiles` controls the log folder and retention. When there are file/task results, Excel contains an `Overview` worksheet and adds `Errors` only when file/task errors occur. No Excel file is created when there is nothing to upload and no file/task errors, including when all source files are excluded. A separate `System errors log.json` contains only system errors, such as configuration or reporting failures, and is created only when those errors occur. File results are not duplicated in JSON. `DeleteLogsAfterDays: 0` keeps logs indefinitely.
- **Overview:** Every file upload and file/task error is recorded with `DateTime`, `ComputerName`, task name, `Type`, `FileName`, `FullName`, `CreationTime`, `LastWriteTime`, `Length` (bytes), `Action`, destination and upload/copy/removal status. File dates are captured before upload and remain available after removal; unavailable metadata is blank. `Action` lists only completed operations, such as `Uploaded, Copied, Removed`. Failed operations appear in `Error` with their `Stage`, while earlier successful actions remain visible.
- **Errors:** A filtered diagnostic view of failed file/task results, including conflicts, copy/removal failures, folder access errors and remote worker failures. It includes the path, error, stage, target, error ID, exception type, script stack trace and source position when available. These failures also remain in `Overview`. When no file/task errors occur, the `Errors` worksheet is not created. System/configuration/reporting failures remain in the separate JSON log.
- **Email:** `Settings.SendMail.When` supports `Never`, `Always`, `OnError`, and `OnErrorOrAction`. To enable it, configure the sender, recipients, SMTP server, and `AssemblyPath` values. MailKit and MimeKit are .NET libraries, not PowerShell modules; compatible DLLs and their dependencies must be installed separately. Leave email disabled until these are available. The email HTML is also saved when file logging is enabled.
- **Windows event log:** Set `Settings.SaveInEventLog.Save` to `true` and specify `LogName`. Ask an administrator to register the event source named by `Settings.ScriptName` if the execution account cannot create it.

## Remote Sources

For a file or folder on another computer, set `ComputerName` to that computer's name and use a path local to it. That computer needs PowerShell 7, `Microsoft.Graph.Authentication`, and access to the certificate/private key. Install the module there using the command above under the executing account, or install it for all users.

PowerShell remoting must be configured and accessible. `PSSessionConfiguration` defaults to `PowerShell.7`. Prefer an archive folder local to the remote computer; accessing a network archive from a remote session can require additional authentication setup.

For network shares, use `ComputerName: null` and UNC paths such as `\\server\share\Reports`. Files are then uploaded by the computer running the command, without remoting.

## Important Behavior

- Existing SharePoint files cause an error unless `SharePoint.Overwrite` is `true`. On a conflict, the source is neither copied nor removed. Missing destination folders are created.
- Source files you keep are attempted again on the next run. Existing destinations then cause errors unless you enable overwriting. Avoid matching destination filenames across tasks unless overwriting is intended.
- Files must be finished writing before upload. Do not run overlapping instances against the same source files or folders.
- Hidden and system files are included unless excluded with `Exclude.Attributes`. Folder links are not followed, and linked source folders or files are rejected. Empty folders are not uploaded; local folders are never removed.
- Each selected folder or explicit file is one job. `MaxConcurrent.JobsTotal` limits simultaneous jobs; `JobsPerComputer` limits them per computer or source file server. Files within each selected folder are processed one at a time.