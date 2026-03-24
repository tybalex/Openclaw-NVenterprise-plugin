---
name: onedrive-upload
description: Upload local files to Microsoft OneDrive or list files in OneDrive folders. Supports sharing uploaded files with specific people (read-only) or via sharing links. Use when the user asks to save, upload, sync, or store files to OneDrive, or browse their OneDrive contents.
---

# OneDrive Upload

## When to Use

Use this skill when the user asks to:
- Upload a file to OneDrive
- Save / store / sync a local file to OneDrive
- Put a file on their Microsoft cloud storage
- Share a file with specific people via OneDrive
- Create a sharing link for a file
- List or browse files in their OneDrive
- Check what's in a OneDrive folder

## Tools

### 1. Upload a file: `onedrive` (action: `upload_file`)

Uploads a local file to the user's Microsoft OneDrive, with optional sharing.

#### Parameters

- **file_path** (REQUIRED, string): Absolute or relative path to the local file. Supports `~` for home directory.
- **destination_path** (optional, string): Where to place the file in OneDrive (e.g. `"Documents/report.pdf"`). If omitted, the file goes to the root with its original name.
- **share_with** (optional, array of strings): Email addresses to grant **read-only** access. Each recipient gets an invitation email.
- **share_link_type** (optional, string): Create a read-only sharing link. Values:
  - `"anonymous"` — anyone with the link can view (no sign-in required)
  - `"organization"` — anyone in the same org can view (sign-in required)
- **share_message** (optional, string): Message to include in the invitation email sent to `share_with` recipients.

#### How to call

**Upload only (no sharing):**

```json
{
  "file_path": "/Users/username/Documents/report.pdf",
  "destination_path": "Documents/Reports/report.pdf"
}
```

**Upload and share with specific people:**

```json
{
  "file_path": "/Users/username/Documents/report.pdf",
  "destination_path": "Documents/Reports/report.pdf",
  "share_with": ["alice@example.com", "bob@example.com"],
  "share_message": "Here's the Q4 report for your review."
}
```

**Upload and create a sharing link:**

```json
{
  "file_path": "/Users/username/slides.pptx",
  "destination_path": "Presentations/slides.pptx",
  "share_link_type": "organization"
}
```

**Upload with both individual sharing and a link:**

```json
{
  "file_path": "/path/to/file.pdf",
  "share_with": ["alice@example.com"],
  "share_link_type": "anonymous"
}
```

#### Examples

| User says | Parameters to pass |
|---|---|
| "Upload report.pdf to my OneDrive" | `file_path` only |
| "Upload slides.pptx and share with alice@nvidia.com" | `file_path` + `share_with: ["alice@nvidia.com"]` |
| "Put this file on OneDrive and give me a link anyone can view" | `file_path` + `share_link_type: "anonymous"` |
| "Save this doc to OneDrive Documents and share with the team, tell them it's the final draft" | `file_path` + `destination_path` + `share_with` + `share_message` |
| "Upload and create an org-only sharing link" | `file_path` + `share_link_type: "organization"` |

### 2. List files: `onedrive` (action: `list_files`)

Lists files and folders in a OneDrive directory.

#### Parameters

- **folder_path** (optional, string): OneDrive folder path to list (e.g. `"Documents"`, `"Projects/2024"`). If omitted, lists the root directory.

#### How to call

```json
{
  "folder_path": "Documents"
}
```

#### Examples

| User says | folder_path |
|---|---|
| "What files do I have on OneDrive?" | (omit — lists root) |
| "Show me what's in my Documents folder on OneDrive" | `"Documents"` |
| "List the files in Projects/Q4" | `"Projects/Q4"` |

## Important Notes

- **Always pass `file_path`** for upload. Never call `onedrive` (action: `upload_file`) with empty arguments `{}`.
- Files under 4 MB use simple upload; larger files automatically use chunked upload sessions.
- The tool reuses the Microsoft Graph authentication. Requires NVIDIA SSO authentication.
- Any file type is supported (documents, images, archives, etc.).
- Sharing is applied **after** the upload completes. If sharing fails, the file is still uploaded successfully.
- `share_with` sends email invitations — recipients must have a Microsoft account or be guest users in the org.
- `share_link_type: "anonymous"` may be blocked by the organization's sharing policy. If so, fall back to `"organization"`.

## Output

- **Upload:** Returns success confirmation with the OneDrive path, web URL, and sharing results.
- **List:** Returns a formatted list of files and folders with sizes and modification dates.
