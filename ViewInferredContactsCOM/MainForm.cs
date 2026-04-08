using System;
using System.Drawing;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace ViewInferredContactsCOM;

public class MainForm : Form
{
    private const string PROP_BASE =
        "http://schemas.microsoft.com/mapi/string/{00062008-0000-0000-C000-000000000046}";
    private const string BUG_REPORT_TO = "sedsats@microsoft.com";

    // Controls
    private WebBrowser browser;
    private ToolStripStatusLabel statusLabel;
    private Button btnReport;

    // Extracted data (retained for bug reports)
    private string emailSubject, emailFrom, emailDate;
    private string rawIsM2H, rawHasTask, rawContactInfo, rawFTSLM, rawClassifier;
    private dynamic currentOutlookApp;
    private dynamic currentMailItem;
    private readonly StringBuilder debugLog = new();

    // ───────────────────────────── UI Setup ─────────────────────────────

    public MainForm()
    {
        Text = "View Inferred Contacts (COM)";
        Size = new Size(580, 780);
        MinimumSize = new Size(420, 500);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9);

        // ── Top button bar ──
        var topPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 48,
            Padding = new Padding(8, 8, 8, 0),
            BackColor = Color.White
        };

        var btnRead = new Button
        {
            Text = "\u2709  Read Selected Email",
            Height = 32,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            Cursor = Cursors.Hand
        };
        btnRead.FlatAppearance.BorderSize = 0;
        btnRead.Click += (_, _) => ReadSelectedEmail();

        btnReport = new Button
        {
            Text = "\u2709  Report Extraction Bug",
            Height = 32,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(232, 232, 232),
            Cursor = Cursors.Hand,
            Enabled = false
        };
        btnReport.FlatAppearance.BorderSize = 0;
        btnReport.Click += (_, _) => ShowBugReportDialog();

        topPanel.Controls.Add(btnRead);
        topPanel.Controls.Add(btnReport);

        // ── WebBrowser for results ──
        browser = new WebBrowser
        {
            Dock = DockStyle.Fill,
            IsWebBrowserContextMenuEnabled = false,
            AllowWebBrowserDrop = false,
            ScriptErrorsSuppressed = true
        };

        // ── Status bar ──
        var statusBar = new StatusStrip();
        statusLabel = new ToolStripStatusLabel(
            "Ready \u2014 Select an email in Outlook and click 'Read Selected Email'");
        statusBar.Items.Add(statusLabel);

        // Add controls (Fill first so it takes remaining space)
        SuspendLayout();
        Controls.Add(browser);
        Controls.Add(topPanel);
        Controls.Add(statusBar);
        ResumeLayout(true);

        ShowWelcomePage();
    }

    // ───────────────────────── HTML Rendering ───────────────────────────

    private static readonly string CssStyle = @"
body { font-family:'Segoe UI',Tahoma,sans-serif; padding:20px; background:#f9f9f9;
       color:#333; font-size:13px; line-height:1.5; margin:0; }
h2   { color:#0078d4; font-size:15px; border-bottom:2px solid #0078d4;
       padding-bottom:4px; margin:20px 0 10px 0; }
h3   { color:#323130; font-size:13px; margin:12px 0 6px 0; }
table { width:100%; border-collapse:collapse; margin:8px 0; }
th,td { padding:6px 8px; text-align:left; border-bottom:1px solid #edebe9; }
th    { background:#f3f2f1; font-weight:600; color:#323130; }
td    { background:#fff; }
pre   { background:#f3f2f1; padding:10px; border:1px solid #e1dfdd; font-size:11px;
        font-family:Consolas,'Courier New',monospace; white-space:pre-wrap;
        word-wrap:break-word; overflow:auto; max-height:300px; }
.ns   { color:#a19f9d; font-style:italic; }
.il   { color:#605e5c; font-weight:600; display:inline-block; min-width:55px; }
.desc { font-size:11px; color:#605e5c; margin:4px 0 12px 0; }
.hi   { color:#107c10; font-weight:600; }
.tog  { cursor:pointer; color:#0078d4; user-select:none; }
.tog:hover { text-decoration:underline; }
";

    private void ShowPage(string bodyHtml)
    {
        browser.DocumentText =
            $"<!DOCTYPE html><html><head>" +
            $"<meta http-equiv='X-UA-Compatible' content='IE=edge'>" +
            $"<style>{CssStyle}</style></head><body>{bodyHtml}</body></html>";
    }

    private void ShowWelcomePage()
    {
        ShowPage(@"
<h2>View Inferred Contacts (COM)</h2>
<p>This tool reads entity extraction properties stamped on emails by the Exchange pipeline,
using <strong>local COM/MAPI</strong> instead of EWS.</p>

<h3>Properties Extracted</h3>
<table>
<tr><th>Property</th><th>Description</th></tr>
<tr><td>EntityExtraction/IsM2H</td><td>Machine-to-Human indicator &mdash; should extraction run?</td></tr>
<tr><td>EntityExtraction/HasTask</td><td>Does the email contain actionable tasks?</td></tr>
<tr><td>EntityExtraction/ContactInformation</td><td>Extracted contact data (shown as vCard)</td></tr>
<tr><td>EntityExtraction/ContactInformationFTSLM</td><td>Fine-tuned supervised learning model output</td></tr>
<tr><td>EntityExtraction/ScalableExtractionClassifierRaw1.0</td><td>Pre-filter classifier scores</td></tr>
</table>

<p class='desc'>Select an email in <strong>Outlook Win32 Desktop</strong> and click
<strong>Read Selected Email</strong> to extract data.</p>
<p class='desc'>This tool does <strong>not</strong> use EWS. All data is read locally via
Outlook COM PropertyAccessor.</p>
");
    }

    private void SetStatus(string text) => statusLabel.Text = text;

    private void Log(string msg) =>
        debugLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {msg}");

    // ──────────────────────── Core COM Logic ────────────────────────────

    private void ReadSelectedEmail()
    {
        debugLog.Clear();
        SetStatus("Reading selected email...");
        Cursor = Cursors.WaitCursor;

        try
        {
            Log("Connecting to Outlook via COM...");
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null) { ShowNotWin32Error(); return; }

            currentOutlookApp = Activator.CreateInstance(outlookType);
            Log($"Connected. Outlook version: {currentOutlookApp.Version}");

            dynamic explorer = currentOutlookApp.ActiveExplorer();
            if (explorer == null)
            {
                SetStatus("Error: No active Outlook window.");
                ShowPage("<h2>Error</h2><p>No active Outlook explorer window found. " +
                         "Please open your Inbox in Outlook.</p>");
                return;
            }

            dynamic selection = explorer.Selection;
            if (selection.Count == 0)
            {
                SetStatus("Error: No email selected.");
                ShowPage("<h2>Error</h2><p>No email is selected. " +
                         "Please select an email in Outlook and try again.</p>");
                return;
            }

            currentMailItem = selection.Item(1);
            Log($"Selected item: {currentMailItem.Subject}");

            // Read email metadata
            try { emailSubject = currentMailItem.Subject; }
            catch { emailSubject = "(unknown)"; }

            try
            {
                string senderType = currentMailItem.SenderEmailType;
                if (senderType == "EX")
                {
                    try
                    {
                        string smtp = currentMailItem.Sender
                            .GetExchangeUser()?.PrimarySmtpAddress;
                        emailFrom = $"{currentMailItem.SenderName} <{smtp}>";
                    }
                    catch
                    {
                        emailFrom = currentMailItem.SenderName;
                    }
                }
                else
                {
                    emailFrom = $"{currentMailItem.SenderName} " +
                                $"<{currentMailItem.SenderEmailAddress}>";
                }
            }
            catch { emailFrom = "(unknown)"; }

            try { emailDate = currentMailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"); }
            catch { emailDate = "(unknown)"; }

            // Read properties via PropertyAccessor
            dynamic pa = currentMailItem.PropertyAccessor;

            rawIsM2H      = ReadProperty(pa, "EntityExtraction/IsM2H");
            rawHasTask     = ReadProperty(pa, "EntityExtraction/HasTask");
            rawContactInfo = ReadProperty(pa, "EntityExtraction/ContactInformation");
            rawFTSLM       = ReadProperty(pa, "EntityExtraction/ContactInformationFTSLM");
            rawClassifier  = ReadProperty(pa, "EntityExtraction/ScalableExtractionClassifierRaw1.0");

            ShowPage(BuildResultsHtml());
            btnReport.Enabled = true;
            SetStatus($"Data extracted from: {emailSubject}");
        }
        catch (COMException ex)
        {
            Log($"COM Error: {ex.Message} (0x{ex.ErrorCode:X8})");
            SetStatus("COM Error \u2014 see details");
            ShowPage($"<h2>COM Error</h2><p>{Enc(ex.Message)}</p>" +
                     $"<p>Error code: 0x{ex.ErrorCode:X8}</p>" +
                     "<p>Make sure Outlook Win32 Desktop is running and an email is selected.</p>");
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
            SetStatus("Error \u2014 see details");
            ShowPage($"<h2>Error</h2><p>{Enc(ex.Message)}</p>");
        }
        finally
        {
            Cursor = Cursors.Default;
        }
    }

    private string ReadProperty(dynamic pa, string propertyName)
    {
        string schema = $"{PROP_BASE}/{propertyName}";
        try
        {
            object val = pa.GetProperty(schema);
            string result = val?.ToString();
            Log($"{propertyName}: {Truncate(result, 120)}");
            return result;
        }
        catch (COMException)
        {
            Log($"{propertyName}: (not stamped)");
            return null;
        }
    }

    // ──────────────────── HTML Result Builder ───────────────────────────

    private string BuildResultsHtml()
    {
        var sb = new StringBuilder();

        // Email info
        sb.Append("<h2>Email Information</h2><div>");
        sb.Append($"<p><span class='il'>Subject:</span> {Enc(emailSubject)}</p>");
        sb.Append($"<p><span class='il'>From:</span> {Enc(emailFrom)}</p>");
        sb.Append($"<p><span class='il'>Date:</span> {Enc(emailDate)}</p>");
        sb.Append("</div>");

        // Pipeline properties
        sb.Append("<h2>Pipeline Properties</h2><table>");
        sb.Append($"<tr><td class='il'>IsM2H</td><td>{Fmt(rawIsM2H)}</td></tr>");
        sb.Append($"<tr><td class='il'>HasTask</td><td>{Fmt(rawHasTask)}</td></tr>");
        sb.Append("</table>");
        sb.Append("<p class='desc'><b>IsM2H:</b> Machine-to-Human indicator. " +
                  "<b>HasTask:</b> Whether the email contains actionable tasks.</p>");

        // Classifier
        sb.Append("<h2>Classifier Scores</h2>");
        sb.Append(FormatClassifierAsHtml(rawClassifier));
        sb.Append("<p class='desc'>Pre-filter scores. If score &gt; threshold, " +
                  "the corresponding extraction pipeline runs.</p>");

        // Contact Information
        sb.Append("<h2>Contact Information</h2>");
        sb.Append(FormatContactInfoAsHtml(rawContactInfo));

        // FTSLM
        sb.Append("<h2>ContactInformationFTSLM</h2>");
        if (rawFTSLM != null)
            sb.Append(FormatContactInfoAsHtml(rawFTSLM));
        else
            sb.Append("<p class='ns'>(not stamped on this email)</p>");
        sb.Append("<p class='desc'>Fine-tuned supervised learning model output " +
                  "for extraction testing.</p>");

        // Debug log (collapsed)
        sb.Append("<h2 class='tog' onclick=\"var d=document.getElementById('dbg');" +
                  "d.style.display=d.style.display==='none'?'block':'none';" +
                  "this.innerText=d.style.display==='none'?" +
                  "'Show Debug Log \u25B6':'Hide Debug Log \u25BC';\">" +
                  "Show Debug Log \u25B6</h2>");
        sb.Append($"<pre id='dbg' style='display:none'>{Enc(debugLog.ToString())}</pre>");

        return sb.ToString();
    }

    private static string Fmt(string val) =>
        val == null ? "<span class='ns'>(not stamped)</span>" : Enc(val);

    // ──────────────────── Classifier Formatting ─────────────────────────

    private string FormatClassifierAsHtml(string json)
    {
        if (json == null)
            return "<p class='ns'>(not stamped on this email)</p>";

        try
        {
            using var doc = JsonDocument.Parse(json);
            var responses = doc.RootElement.GetProperty("ClassifierResponses");

            var sb = new StringBuilder();
            sb.Append("<table><thead><tr><th>Entity</th><th>Score</th>" +
                      "<th>Threshold</th></tr></thead><tbody>");

            foreach (var item in responses.EnumerateArray())
            {
                string entity = item.GetProperty("entity").GetString();
                double score = item.GetProperty("score").GetDouble();
                double threshold = item.GetProperty("threshold").GetDouble();
                string cls = score > threshold ? " class='hi'" : "";

                sb.Append($"<tr><td>{Enc(entity)}</td>" +
                          $"<td{cls}>{score:F4}</td>" +
                          $"<td>{threshold:F1}</td></tr>");
            }

            sb.Append("</tbody></table>");

            if (doc.RootElement.TryGetProperty("CorrelationId", out var cid))
                sb.Append($"<p class='desc'>CorrelationId: {Enc(cid.GetString())}</p>");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"<p>Error parsing classifier data: {Enc(ex.Message)}</p>" +
                   $"<pre>{Enc(json)}</pre>";
        }
    }

    // ──────────────── Contact Information Formatting ─────────────────────

    private string FormatContactInfoAsHtml(string json)
    {
        if (json == null)
            return "<p class='ns'>(not stamped on this email)</p>";

        try
        {
            using var doc = JsonDocument.Parse(json);

            if (!doc.RootElement.TryGetProperty("ContactInfo", out var contacts))
                return $"<pre>{Enc(json)}</pre>";

            var sb = new StringBuilder();

            // vCard output
            sb.Append("<h3>vCard Output</h3><pre>");

            foreach (var contact in contacts.EnumerateArray())
            {
                string displayName = JsonPath(contact, "Name", "DisplayName") ?? "Unknown";
                string first  = JsonPath(contact, "Name", "AdditionalInfo", "First") ?? "";
                string middle = JsonPath(contact, "Name", "AdditionalInfo", "Middle") ?? "";
                string last   = JsonPath(contact, "Name", "AdditionalInfo", "Last") ?? "";
                string nick   = JsonPath(contact, "Name", "AdditionalInfo", "Nickname") ?? "";
                string title  = JsonPath(contact, "Position", "JobTitle") ?? "";

                sb.AppendLine("BEGIN:VCARD");
                sb.AppendLine("VERSION:3.0");
                sb.AppendLine($"FN:{displayName}");
                sb.AppendLine($"N:{last};{first};{middle};;");
                if (!string.IsNullOrEmpty(nick))  sb.AppendLine($"NICKNAME:{nick}");
                if (!string.IsNullOrEmpty(title)) sb.AppendLine($"TITLE:{title}");

                if (contact.TryGetProperty("Emails", out var emails))
                    foreach (var e in emails.EnumerateArray())
                        if (e.TryGetProperty("Address", out var addr))
                            sb.AppendLine($"EMAIL:{addr.GetString()}");

                if (contact.TryGetProperty("Phones", out var phones))
                    foreach (var p in phones.EnumerateArray())
                        if (p.TryGetProperty("Number", out var num))
                            sb.AppendLine($"TEL:{num.GetString()}");

                if (contact.TryGetProperty("Addresses", out var addrs))
                    foreach (var a in addrs.EnumerateArray())
                        if (a.TryGetProperty("AddressValue", out var av))
                            sb.AppendLine($"ADR:{av.GetString()}");

                sb.AppendLine("END:VCARD");
                sb.AppendLine();
            }

            sb.Append("</pre>");

            // Extraction metadata
            if (doc.RootElement.TryGetProperty("ExtractionTimeUTC", out var extTime))
                sb.Append($"<p class='desc'>Extraction Time (UTC): " +
                          $"{Enc(extTime.GetString())}</p>");
            if (doc.RootElement.TryGetProperty("CorrelationTrail", out var trail))
                sb.Append($"<p class='desc'>CorrelationTrail: " +
                          $"{Enc(trail.GetString())}</p>");

            // Raw JSON
            sb.Append("<h3>Raw JSON</h3>");
            sb.Append($"<pre>{Enc(json)}</pre>");

            return sb.ToString();
        }
        catch (Exception ex)
        {
            return $"<p>Error parsing contact info: {Enc(ex.Message)}</p>" +
                   $"<pre>{Enc(json)}</pre>";
        }
    }

    private static string JsonPath(JsonElement root, params string[] path)
    {
        var current = root;
        foreach (var key in path)
            if (!current.TryGetProperty(key, out current))
                return null;
        return current.ValueKind == JsonValueKind.String
            ? current.GetString()
            : current.ToString();
    }

    // ───────────────────────── Bug Report ────────────────────────────────

    private void ShowBugReportDialog()
    {
        using var dialog = new Form
        {
            Text = "Report Extraction Bug",
            Size = new Size(460, 300),
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            StartPosition = FormStartPosition.CenterParent,
            Font = new Font("Segoe UI", 9)
        };

        var lbl = new Label
        {
            Text = "Can you describe what's missing or wrong?",
            Location = new Point(15, 15),
            AutoSize = true
        };
        var txtDesc = new TextBox
        {
            Location = new Point(15, 40),
            Size = new Size(415, 100),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        var chkInclude = new CheckBox
        {
            Text = "Include the original email content in the report",
            Location = new Point(15, 150),
            AutoSize = true
        };

        var btnSend = new Button
        {
            Text = "Send Report",
            Location = new Point(15, 190),
            Size = new Size(110, 32),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White
        };
        btnSend.FlatAppearance.BorderSize = 0;

        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(135, 190),
            Size = new Size(80, 32),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(232, 232, 232)
        };
        btnCancel.FlatAppearance.BorderSize = 0;

        btnSend.Click += (_, _) =>
        {
            SendBugReport(txtDesc.Text, chkInclude.Checked);
            dialog.Close();
        };
        btnCancel.Click += (_, _) => dialog.Close();

        dialog.Controls.AddRange(new Control[] { lbl, txtDesc, chkInclude, btnSend, btnCancel });
        dialog.ShowDialog(this);
    }

    private void SendBugReport(string description, bool includeEmail)
    {
        try
        {
            if (currentOutlookApp == null)
            {
                MessageBox.Show("Not connected to Outlook.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            dynamic mail = currentOutlookApp.CreateItem(0); // olMailItem
            mail.To = BUG_REPORT_TO;
            mail.Subject = "Bug from plugin (COM)";

            string originalContent = "";
            if (includeEmail && currentMailItem != null)
            {
                try { originalContent = currentMailItem.HTMLBody; }
                catch { originalContent = "Could not retrieve original email content."; }
            }

            string body = $@"
<p><strong>Bug Report (COM Plugin)</strong></p>
<br>
<p><strong>Description:</strong> {Enc(description)}</p>
<br>
<p><strong>Email Subject:</strong> {Enc(emailSubject)}</p>
<p><strong>Email From:</strong> {Enc(emailFrom)}</p>
<p><strong>Email Date:</strong> {Enc(emailDate)}</p>
<br>
<p><strong>IsM2H:</strong> {Enc(rawIsM2H ?? "(not stamped)")}</p>
<p><strong>HasTask:</strong> {Enc(rawHasTask ?? "(not stamped)")}</p>
<p><strong>ScalableExtractionClassifierRaw1.0:</strong> {Enc(rawClassifier ?? "(not stamped)")}</p>
<p><strong>ContactInformation:</strong> {Enc(rawContactInfo ?? "(not stamped)")}</p>
<p><strong>ContactInformationFTSLM:</strong> {Enc(rawFTSLM ?? "(not stamped)")}</p>
<br>
{(includeEmail && !string.IsNullOrEmpty(originalContent)
    ? $@"<p><strong>Original Email Content:</strong></p>
         <pre>=== ORIGINAL EMAIL BEGINS ===
{originalContent}
=== ORIGINAL EMAIL ENDS ===</pre>"
    : "")}
";
            mail.HTMLBody = body;
            mail.Display(); // Opens compose window for review before sending
            SetStatus("Bug report email created \u2014 review and send.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error creating bug report: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ──────────────────────── Error Pages ────────────────────────────────

    private void ShowNotWin32Error()
    {
        ShowPage(@"
<h2>Outlook Win32 Desktop Required</h2>
<p>This tool requires <strong>Microsoft Outlook Win32 Desktop</strong> (classic Outlook)
with COM support.</p>
<p>It will <strong>not</strong> work with:</p>
<ul>
<li>New Outlook (Monarch)</li>
<li>Outlook on the web</li>
<li>Outlook without COM support</li>
</ul>
<p>Please switch to Outlook Win32 Desktop and try again.</p>
");
        SetStatus("Error: Outlook Win32 Desktop not detected.");
    }

    // ──────────────────────── Helpers ────────────────────────────────────

    private static string Enc(string text) =>
        string.IsNullOrEmpty(text) ? "" : WebUtility.HtmlEncode(text);

    private static string Truncate(string s, int max) =>
        s == null ? "(null)" : s.Length <= max ? s : s[..max] + "...";
}
