# emailcli — Outlook from the command line

A lightweight PowerShell toolbox to read, manage and compose Outlook email, calendar events and contacts from the terminal.

Author: Tom Villani, Ph.D.  
Module file: emailcli.psm1

Repository contents
- `emailcli.psm1` – main PowerShell module implementation.
- `emailcli.psd1` – module manifest.
- `.emailcli.json.example` – example configuration file; copy to `%USERPROFILE%\.emailcli.json`.
- `LICENSE` – MIT license.
- `readme.md` – this documentation.


Features
- Query your Outlook Inbox with many useful filters (unread, replied, flagged, from, subject, category, attachments, importance, date ranges).
- Read messages in the terminal with cleaned-up body text (SafeLinks decoding, tracking-link simplification, HTML entity cleanup).
- Save attachments (skips small inline images by default).
- Reply / forward / compose new messages (opens system editor when body not provided).
- Archive messages (move to Outlook Archive folder).
- Search Archive folder with mandatory filters to avoid huge result sets.
- Interactive mode for reviewing messages and taking actions (reply, forward, archive, flag, categorize, save attachments, add contacts).
- Calendar helper functions (day/week view, create events).
- Contact helpers (list/search, add, sync from Sent items).
- Configuration via `~/.emailcli.json` for sensible defaults.

Requirements
- Windows with Outlook installed and configured (module uses Outlook COM objects).
- PowerShell (tested in Windows PowerShell; running in PowerShell Core on Windows may work but COM support and Outlook availability are required).
- Permissions to access Outlook via the local COM API (run as your user with Outlook configured).

Installation
1. Save the provided `emailcli.psm1` and `emailcli.psd1` somewhere convenient (for example: `%USERPROFILE%\Documents\WindowsPowerShell\Modules\emailcli\`).
   - If you place them in a folder named `emailcli` under one of your PowerShell module paths, you can simply `Import-Module emailcli`.
   - Or import directly from a file:
     ```powershell
     Import-Module C:\path\to\emailcli.psm1
     ```
2. Optionally add `Import-Module emailcli` to your PowerShell profile so the functions/aliases are available in each session:
   ```powershell
   # Add to your profile
   "Import-Module emailcli" | Out-File -FilePath $PROFILE -Append -Encoding utf8
   ```


Configuration (~/.emailcli.json)
An example config is provided as `.emailcli.json.example` in this repo; copy it to `%USERPROFILE%\.emailcli.json` and edit as needed.


The module looks for a JSON config file at `%USERPROFILE%\.emailcli.json` and uses values as defaults when parameters aren't provided.

Example `~/.emailcli.json`:
```json
{
  "Email": {
    "DaysBack": 7,
    "IncludeRead": false,
    "Limit": 50,
    "Compact": true,
    "KeepLinks": false,
    "Oldest": false
  },
  "Interactive": {
    "DaysBack": 7,
    "PreviewLength": 600,
    "Limit": 0,
    "IncludeRead": false,
    "KeepLinks": false,
    "Oldest": false
  },
  "Archive": {
    "DaysBack": 90,
    "Limit": 50,
    "Compact": false
  },
  "Calendar": {
    "Days": 7,
    "Compact": true,
    "WorkWeek": false
  },
  "ReadMessage": {
    "AttachmentPath": ".",
    "SmallImageThreshold": 100,
    "IncludeSmallImages": false
  },
  "SendEmail": {
    "Importance": "Normal"
  },
  "Contacts": {
    "DaysBack": 30
  }
}
```

Use Set-ConfigValue / Get-ConfigValue to update/read config from PowerShell:
```powershell
# Set a config value (persists to ~/.emailcli.json)
Set-ConfigValue -Section Interactive -Key PreviewLength -Value 800

# Read a config value
Get-ConfigValue -Section Email -Key DaysBack -Default 7
```


Quick start — examples
- List recent unread email (default 7 days):
  ```powershell
  Get-OutlookInbox
  # alias: inbox
  ```
- List last 30 days from a particular sender:
  ```powershell
  Get-OutlookInbox -From "john.doe" -DaysBack 30
  ```
- Show compact view with 10 most recent messages:
  ```powershell
  Get-OutlookInbox -Compact -Limit 10
  ```
- Read a message from the last Get-OutlookInbox output:
  ```powershell
  Read-OutlookMessage -Index 0
  # alias: read-email
  ```
- Read message in plain text (suitable for piping):
  ```powershell
  Get-OutlookInbox -Plain | Out-File inbox.txt
  Read-OutlookMessage -Index 0 -Plain | Out-File message0.txt
  ```
- Save attachments from a message:
  ```powershell
  Read-OutlookMessage -Index 4 -SaveAttachments -AttachmentPath "C:\Temp\email-attachments"
  ```
- Reply to message (opens editor if no -Body provided):
  ```powershell
  Send-OutlookReply -Index 2
  # alias: reply-email
  ```
- Forward:
  ```powershell
  Send-ForwardOutlookEmail -Index 3 -To "recipient@example.com"
  # alias: forward-email
  ```
- Archive messages by index:
  ```powershell
  Move-OutlookMessageToArchive -Index 0,1 -MarkAsRead
  # alias: archive-email
  ```
- Search Archive (must include at least one of -From, -Subject, or -Category):
  ```powershell
  Search-OutlookArchive -From "alice" -DaysBack 180 -Limit 100
  # alias: search-archive
  ```
- Start interactive review mode:
  ```powershell
  Start-OutlookInteractive
  # alias: iinbox
  ```

Calendar
- Today:
  ```powershell
  Get-OutlookCalendarDay
  # alias: calendar-today or calendar-day
  ```
- Week:
  ```powershell
  Get-OutlookCalendarWeek -StartDate (Get-Date) -Days 7
  # alias: calendar-week
  ```
- Create event:
  ```powershell
  New-OutlookCalendarEvent -Subject "Team Sync" -Start "2025-12-01 10:00" -Duration 60 -Attendees "a@example.com","b@example.com" -SendInvites
  # alias: new-event
  ```

Contacts
- Add new contact:
  ```powershell
  Add-OutlookContact -Email "john.doe@example.com" -FirstName John -LastName Doe
  # alias: add-contact
  ```
- List/search contacts:
  ```powershell
  Get-OutlookContacts -Search "john"
  # alias: contacts
  ```
- Sync contacts from recent sent replies:
  ```powershell
  Sync-OutlookContactsFromReplies -DaysBack 30 -Interactive
  # alias: sync-contacts
  ```

Helpful commands
- View currently cached inbox messages (from last Get-OutlookInbox): Get-CachedMessages
- View cached archive search results: Get-CachedArchiveMessages
- Clear config cache (force reload of ~/.emailcli.json): Reset-EmailCliConfig
- Reset category color cache (refresh Outlook categories): Reset-CategoryColorCache

Aliases
The module sets a number of convenience aliases:
- inbox -> Get-OutlookInbox
- read-email -> Read-OutlookMessage
- reply-email -> Send-OutlookReply
- forward-email -> Send-ForwardOutlookEmail
- iinbox -> Start-OutlookInteractive
- email-categories -> Get-OutlookCategories
- archive-email -> Move-OutlookMessageToArchive
- search-archive -> Search-OutlookArchive
- calendar-today/day/week -> Get-OutlookCalendarDay / Get-OutlookCalendarWeek
- new-event -> New-OutlookCalendarEvent
- add-contact -> Add-OutlookContact
- contacts -> Get-OutlookContacts
- sync-contacts -> Sync-OutlookContactsFromReplies
- flag-email -> Add-OutlookFlag
- tag-email -> Add-OutlookCategory
- refresh-categories -> Reset-CategoryColorCache
- send-email -> Send-OutlookEmail

Exported functions (summary)
- Get-OutlookInbox, Read-OutlookMessage, Send-OutlookReply, Send-ForwardOutlookEmail, Start-OutlookInteractive
- Get-OutlookCategories, Send-OutlookEmail, Move-OutlookMessageToArchive
- Search-OutlookArchive, Get-CachedArchiveMessages, Get-CachedMessages
- Get-OutlookCalendarDay, Get-OutlookCalendarWeek, New-OutlookCalendarEvent
- Add-OutlookContact, Get-OutlookContacts, Sync-OutlookContactsFromReplies
- Add-OutlookFlag, Add-OutlookCategory, Reset-CategoryColorCache
- Configuration helpers: Get-EmailCliConfig, Get-ConfigValue, Set-ConfigValue, Reset-EmailCliConfig
(See module file for full function docstrings and parameter lists.)

Notes, tips & troubleshooting
- Outlook must be installed and configured for your user. The module uses COM automation (New-Object -ComObject Outlook.Application).
- If Outlook prompts for security or blocks programmatic access, adjust your trust/security settings (IT policy may be restricting COM access).
- The module uses the `micro` editor (invoked as `micro <tempfile>`) for interactive composition. If you don't have it, install it or edit the module to run your editor of choice (e.g., `notepad`, `code --wait`, etc.).
- Archive folder detection: the module first tries the Outlook default folder enum for Archive; if that fails it looks for a folder named "Archive". If you have a custom archive path, verify the module finds it.
- Archive search requires at least one of -From, -Subject or -Category to avoid scanning huge numbers of messages.
- The module attempts to decode Microsoft SafeLinks and simplify tracking URLs. If a particular tracking domain isn't handled, you can extend `Clean-EmailBody` in the module.
- If you change Outlook categories while a session is open, run `Reset-CategoryColorCache` (or `refresh-categories` alias) so new categories will be displayed with colors.
- Use `Reset-EmailCliConfig` if you edit `~/.emailcli.json` and want the running session to pick up the new configuration.
- If you see COM object errors or "Failed to access Outlook" messages, ensure Outlook is installed, accessible, and you have the required permissions.

Security & privacy
- This module accesses your local Outlook profile and messages. Be mindful of running any code that manipulates or sends email on your behalf.
- When using automation that sends messages, the module prompts for confirmation before sending.

License & attribution
- Copyright (c) 2025 Tom Villani, Ph.D.
- Licensed under the MIT License (see `LICENSE`); use and modify at your own risk.

Support / Contributing
- The module is intended as a handy toolbox and is easy to extend. Edit the `.psm1` to add or tweak behavior (e.g., link-handling, editors, filters).
- If you add improvements, consider keeping documentation updated and adding configuration-driven editor selection (instead of hard-coded `micro`) if you fork or adapt the module.
