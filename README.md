# Upload Files to SharePoint

Uploads files from local folders, UNC shares, or folders on remote Windows computers to SharePoint document libraries. The entry point has one script parameter:

```powershell
& '.\Main.ps1' -ConfigurationJsonFile '.\MyConfig.json'
```

Start with [Example.json](Example.json). Replace the example paths, sites, authentication, and reporting settings before running it. The second example archives and removes source files. There is no preview/WhatIf mode.

## Requirements

- PowerShell 7 on the script host and every remote worker.
- `ImportExcel` on the script host, for Excel reporting.
- `Microsoft.Graph.Authentication` on every computer doing uploads.
- An Entra ID application with certificate-based authentication. Grant application permission `Sites.ReadWrite.All`, or `Sites.Selected` with an explicit **write** grant on each target site, with administrator consent as applicable.
- Install the certificate and its accessible private key for the executing account on each worker computer. The account also needs source read permissions, archive write permissions when copying, and source delete permissions when removing.
- For remote paths: a configured PowerShell 7 remoting endpoint and permission to open a session. For UNC sources, use `ComputerName: null` to run on the script host without a remoting double hop.
- MailKit/MimeKit assemblies and their dependencies when mail is enabled. Provision the Windows event source beforehand if the execution account cannot create it.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

## Configuration

The configuration uses the same `MaxConcurrent`, `Tasks`, `PSSessionConfiguration`, and `Settings` layout as the removal script. The mail, file logging, and event logging settings keep the same property names. Upload tasks use `Folders`, not the removal-specific `Files`, `OlderThan`, `Exclude`, or `RemoveEmptyFolders` options. The `?` section in the example describes each upload setting.

Each task contains:

- `ComputerName`: `localhost`, a remote computer name, or `null` for UNC sources.
- `Folders`: an array of absolute paths or `{ "Name": "Reports", "Path": "C:\\Exports" }` objects.
- `Recurse`: required Boolean. With `true`, relative subfolders are preserved in SharePoint and the archive. Source root folder names are not included in the destination path.
- `SharePoint`: site URL, document library name, optional library folder path, and certificate authentication settings. `MaxRetries` defaults to 3 and `ChunkSizeMB` to 5.
- `AfterUpload`: optional copy and removal actions, shown below.

| Outcome after successful upload | RemoveFile | CopyToFolder |
| --- | --- | --- |
| Keep the source | `false` | `null` |
| Remove the source | `true` | `null` |
| Copy elsewhere and keep the source | `false` | Absolute archive folder |
| Archive, then remove the source | `true` | Absolute archive folder |

```json
"AfterUpload": {
  "RemoveFile": true,
  "CopyToFolder": "C:\\Archive\\Reports"
}
```

Omitting `AfterUpload` keeps source files without copying. If copying is enabled, it must succeed before removal. Existing archive files are **not overwritten**: collisions are errors, and the source remains. Missing archive folders are created. Archive paths must be outside every source folder on that computer. Source folders must not overlap.

Authentication fields accept `ENV:VARIABLE_NAME`, resolved by the uploader on the worker computer. Folder paths, archive paths, and supported reporting fields resolve their `ENV:` references on the script host. SharePoint URLs and library/folder names are literal strings. All tasks in one configuration must specify identical authentication settings, because local parallel workers share the Graph authentication context. Do not switch Graph identities from other commands in the same PowerShell process during a run.

`MaxConcurrent.JobsTotal` limits simultaneous folder jobs. `JobsPerComputer` limits workers per target computer, or per source file server for UNC paths. Files inside a folder are uploaded sequentially. Remote archive access through UNC paths may need explicitly configured delegation; prefer archive paths local to that worker.

## Upload Behavior

The bundled [UploadToSharePoint.ps1](UploadToSharePoint.ps1) is copied unchanged from `Permission-matrix-set-permissions/Scripts/Operations/UploadToSharePoint.ps1`. It creates missing SharePoint folders, overwrites existing destination files, and uses chunked upload sessions for files larger than 4 MB. Version history depends on the destination library settings.

The worker requires a returned SharePoint item ID before post-upload actions. It holds the source open against writes/deletion during upload and copying, then checks its size and modification time before removal. Producers should finish writing files before placing them in the upload folder. This is not an atomic transaction across SharePoint and the filesystem: do not replace source files concurrently with cleanup or run overlapping script instances.

Hidden/system files are included. Directory links are not traversed; linked source roots and linked files are rejected. Empty folders are not created in SharePoint, and local folders are never removed.

A file error does not prevent other files from being processed. An unreadable or missing source is reported as an error. Failed uploads never trigger copying or removal; failed copying leaves the uploaded SharePoint file and original source in place. Whole remote jobs are not automatically retried, to avoid repeating post-upload actions. Graph requests retain the original uploader's retry behavior.

Multiple source folders can target the same SharePoint folder. Ensure their relative file paths do not collide unless overwriting is intended. There is no upload-history database: retained source files upload again on later runs. Keep logs and archives outside upload sources, including paths that alias the same location through a share or junction.

## Reporting

`Settings` preserves `ScriptName`, `SendMail`, `SaveLogFiles`, and `SaveInEventLog` from the removal configuration. `SendMail.When` supports `Never`, `Always`, `OnError`, and `OnErrorOrAction`; an action means a confirmed upload, including one whose later copy/removal failed.

With file logging enabled, the script saves a JSON run log, an Excel Overview sheet when results exist, and the email HTML when mailing. Each file result contains its source, destination, SharePoint URL, upload status, archive path, removal status, and error. JSON also records configuration and reporting failures. Retention only removes this script/configuration's generated logs in the configured folder, using elapsed days since last modification.

The script outputs per-file results and exits with code 1 if any file, configuration, execution, or reporting error occurs. Success returns normally. Email and event-log errors are included in JSON when it can be written. Failures that happen after a notification has been sent cannot appear retroactively in that notification.

## Files and Tests

Deploy all four PowerShell files together. There is no runtime dependency on sibling script folders. [Functions.ps1](Functions.ps1) contains the configuration and reporting functions alongside unchanged copies of the log-folder, environment-value, parallel execution, MailKit, and event-log helpers from the removal script. Only the folder-upload orchestration and report assembly are new.

```powershell
Invoke-Pester -Path '.\Tests' -Output Detailed
```

Tests use temporary local files and mocked Graph, SMTP, and remoting calls. A live SharePoint upload and account/certificate permissions must be verified in the target environment before enabling source removal.