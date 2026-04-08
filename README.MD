# View Inferred Contacts

A debugging toolkit for inspecting **entity extraction properties** that the Exchange pipeline stamps on emails. Includes two implementations:

1. **Web Add-in** (EWS) -- runs as an Outlook task pane via Office.js
2. **COM Desktop Tool** (.NET 8 / Windows Forms) -- reads properties locally via MAPI, built as an EWS replacement ahead of the [Oct 2026 deprecation](https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/deprecation-of-ews-exchange-online)

---

## Properties Inspected

Both tools read properties from the MAPI namespace `{00062008-0000-0000-C000-000000000046}` (PSETID_Common):

| Property | Type | Description |
|---|---|---|
| `EntityExtraction/IsM2H` | Boolean | Machine-to-Human indicator -- should extraction run? |
| `EntityExtraction/HasTask` | Boolean | Does the email contain actionable tasks? |
| `EntityExtraction/ContactInformation` | JSON | Extracted contact data (names, emails, phones, addresses) |
| `EntityExtraction/ContactInformationFTSLM` | JSON | Fine-tuned supervised learning model output |
| `EntityExtraction/ScalableExtractionClassifierRaw1.0` | JSON | Pre-filter classifier scores per entity type |

Contact information is rendered as **vCard V3.0** for easy readability.

---

## Option 1: COM Desktop Tool (Recommended)

A standalone Windows Forms app that connects to a running Outlook Win32 instance via COM. No network calls, no OAuth, works fully offline.

### Requirements

- **Windows 10/11**
- **Outlook Win32 Desktop** (classic Outlook) -- will **not** work with New Outlook or Outlook on the web
- **.NET 8 SDK** -- [download here](https://dotnet.microsoft.com/download/dotnet/8.0)

### How to Run

```bash
cd ViewInferredContactsCOM
dotnet run
```

This opens a small **Windows Forms popup window** that sits alongside your Outlook window:

1. Open **Outlook Win32 Desktop** (the classic/legacy version)
2. Select any email in the mail list
3. In the popup tool, click **"Read Selected Email"**
4. The tool reads the MAPI properties from the selected email and displays the results

### How to Debug

Open the project in Visual Studio or VS Code:

```bash
cd ViewInferredContactsCOM
dotnet build
```

**In Visual Studio:**
- Open `ViewInferredContactsCOM.csproj`
- Press **F5** to launch with debugger attached
- Set breakpoints in `MainForm.cs` (e.g., `ReadSelectedEmail()` on line ~100, or `ReadProperty()`)

**In VS Code:**
- Open the workspace root -- the `.vscode/launch.json` is pre-configured
- Select the debug configuration and press **F5**
- The popup window launches with debugger attached

**From the command line:**
```bash
cd ViewInferredContactsCOM
dotnet run
```

### How It Works

```
Your desktop
+--------------------------+  +----------------------------+
|  Outlook Win32 (classic) |  |  ViewInferredContactsCOM   |
|                          |  |  (popup window)            |
|  [Selected email] -------|->|  Click "Read Selected      |
|                          |  |   Email"                   |
|  COM/MAPI server --------|->|  Reads MAPI properties     |
|                          |  |  via PropertyAccessor      |
|                          |  |                            |
|                          |  |  Displays:                 |
|                          |  |  - IsM2H flag              |
|                          |  |  - HasTask flag            |
|                          |  |  - Contact info (vCard)    |
|                          |  |  - Classifier scores       |
+--------------------------+  +----------------------------+
```

The tool calls `Type.GetTypeFromProgID("Outlook.Application")` to attach to the running Outlook process, then reads the selected email's properties through `MailItem.PropertyAccessor.GetProperty()`.

### Build a Standalone .exe

To share with teammates who don't have the .NET SDK installed:

```bash
cd ViewInferredContactsCOM
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

Output: `bin/Release/net8.0-windows/win-x64/publish/ViewInferredContactsCOM.exe` -- a single file, no dependencies needed.

---

## Option 2: Web Add-in (EWS)

An Outlook task pane add-in that uses Exchange Web Services to fetch entity extraction properties. Shows up as a **"Debug Email"** button in the Outlook ribbon.

> **Note:** EWS is being deprecated in October 2026. Use the COM tool above for a future-proof alternative.

### Requirements

- **Node.js** (v14+)
- **Outlook 2016 or later** (any platform: Win32, Mac, Web)
- **Exchange Web Services** enabled on your Exchange server

### How to Run

```bash
# Install dependencies
npm install

# Start dev server + sideload into Outlook
npm start
```

This will:
1. Generate dev SSL certificates (first run only)
2. Start a webpack dev server on `https://localhost:3000`
3. Sideload the add-in into Outlook
4. A **"Debug Email"** button appears in the Outlook ribbon under the "Debugger" group

### How to Use

1. Open any email in Outlook
2. Click **"Debug Email"** in the ribbon (or find it in the task pane)
3. Click **"Get extracted data"** in the task pane
4. View the extracted properties, formatted contact info, and classifier scores

### How to Debug

**In VS Code:**
- Pre-configured launch configs exist in `.vscode/launch.json` for Edge Chromium and Edge Legacy
- Press **F5** -- this runs `npm start` as a pre-launch task and attaches the debugger
- Set breakpoints in `src/taskpane/taskpane.js`

**Build for production:**
```bash
npm run build
```

Output goes to `dist/` -- the production add-in is hosted at `https://inferredcontacts.z1.web.core.windows.net/`.

---

## Comparison: COM vs. EWS

| | EWS Web Add-in | COM Desktop Tool |
|---|---|---|
| **EWS Deprecation** | Affected (Oct 2026) | Not affected |
| **Network** | Requires Exchange connectivity | Fully local (offline OK) |
| **Auth** | OAuth tokens | Uses running Outlook session |
| **Deployment** | Hosted web app | `dotnet run` or standalone .exe |
| **Platform** | Any Outlook client | Win32 Desktop only |
| **Debugging** | Browser DevTools + VS Code | .NET debugger (VS / VS Code) |

---

## Bug Reporting

Both tools include a **Report Extraction Bug** feature that composes an email to the pipeline team (`sedsats@microsoft.com`) with:
- Your description of the issue
- All extracted property values
- Optionally the original email content

The email opens in Outlook for review before sending.

---

## Project Structure

```
View-Inferred-Contacts/
|-- manifest.xml                    # Outlook web add-in manifest
|-- package.json                    # Node.js dependencies & scripts
|-- webpack.config.js               # Webpack build config
|-- src/taskpane/
|   |-- taskpane.html               # Web add-in UI
|   |-- taskpane.js                 # EWS request logic
|   +-- taskpane.css                # Styling
|-- ViewInferredContactsCOM/
|   |-- ViewInferredContactsCOM.csproj  # .NET 8 project
|   |-- MainForm.cs                     # Windows Forms UI & COM logic
|   +-- Program.cs                      # Entry point
|-- assets/                         # Add-in icons
+-- .vscode/                        # Debug & build configurations
```

## License

MIT

## Contact

[aaanerud@microsoft.com](mailto:aaanerud@microsoft.com)
