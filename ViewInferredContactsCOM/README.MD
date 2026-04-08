# View Inferred Contacts (COM)

A lightweight Windows desktop tool that reads **entity extraction properties** stamped on emails by the Exchange pipeline, using **local COM/MAPI** instead of EWS.

Built as a replacement for the [EWS-based Outlook add-in](https://github.com/Aanerud/View-Inferred-Contacts) ahead of the [EWS deprecation (Oct 2026)](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online).

## Requirements

- **Outlook Win32 Desktop** (classic Outlook) — will not work with New Outlook (Monarch) or Outlook on the web
- **.NET 8 SDK** — [download here](https://dotnet.microsoft.com/download/dotnet/8.0)
- **Windows 10/11**

## Quick Start

```bash
git clone https://github.com/Aanerud/View-Inferred-Contacts-COM.git
cd View-Inferred-Contacts-COM
dotnet run
```

1. Open **Outlook Win32 Desktop** and select an email
2. Click **Read Selected Email** in the tool
3. View the extracted pipeline properties

## Properties Extracted

| Property | Type | Description |
|---|---|---|
| `EntityExtraction/IsM2H` | Boolean | Machine-to-Human indicator — should extraction run? |
| `EntityExtraction/HasTask` | Boolean | Does the email contain actionable tasks? |
| `EntityExtraction/ContactInformation` | JSON | Extracted contact data (displayed as vCard V3.0) |
| `EntityExtraction/ContactInformationFTSLM` | JSON | Fine-tuned supervised learning model output |
| `EntityExtraction/ScalableExtractionClassifierRaw1.0` | JSON | Pre-filter classifier scores per entity type |

All properties are read from the MAPI named property namespace `{00062008-0000-0000-C000-000000000046}` (PSETID_Common) via Outlook's `PropertyAccessor.GetProperty()`.

## How It Works

```
Outlook Win32 (COM Server)
    |
    +-- ActiveExplorer().Selection   --> Get selected email
    |
    +-- MailItem.PropertyAccessor    --> Read MAPI named properties
    |       |
    |       +-- GetProperty("http://schemas.microsoft.com/mapi/string/{GUID}/EntityExtraction/...")
    |
    +-- Application.CreateItem(0)    --> Compose bug report email
```

- **No EWS** — reads directly from the local OST/PST cache via COM
- **No network calls** — works fully offline once emails are cached
- **No OAuth/auth setup** — uses the running Outlook session
- **No install** — just `dotnet run` (or build a standalone .exe)

## Bug Reporting

Click **Report Extraction Bug** to compose an email to the pipeline team (`sedsats@microsoft.com`) with:
- Your description of what's wrong
- All extracted property values
- Optionally the original email content

The email opens in Outlook's compose window for review before sending.

## Build a Standalone .exe

To share with team members who don't have the .NET SDK:

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The output will be in `bin/Release/net8.0-windows/win-x64/publish/ViewInferredContactsCOM.exe` — a single file that runs without any dependencies.

## Why COM Instead of EWS?

| | EWS (old) | COM (this tool) |
|---|---|---|
| **Deprecation** | Retiring Oct 2026 | Not affected |
| **Network** | Requires Exchange connectivity | Fully local |
| **Auth** | OAuth tokens required | Uses running Outlook session |
| **Deployment** | Web add-in (hosted) | Desktop app (no hosting) |
| **Platform** | Any Outlook client | Win32 Desktop only |

## Platform Support

| Platform | Supported |
|---|---|
| Outlook Win32 Desktop (classic) | Yes |
| New Outlook (Monarch) | No — shows error on launch |
| Outlook on the web | No — shows error on launch |
| macOS Outlook | No — Windows only |
