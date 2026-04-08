using System;
using System.Windows.Forms;

namespace ViewInferredContactsCOM;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Verify Outlook COM is registered on this system.
        // New Outlook (Monarch) and Outlook on the web do NOT register COM ProgIDs,
        // so this check ensures we are on a Win32 Desktop Outlook machine.
        var outlookType = Type.GetTypeFromProgID("Outlook.Application");
        if (outlookType == null)
        {
            MessageBox.Show(
                "This tool requires Microsoft Outlook Win32 Desktop (classic Outlook).\n\n" +
                "It will not work with:\n" +
                "  \u2022 New Outlook (Monarch)\n" +
                "  \u2022 Outlook on the web\n" +
                "  \u2022 Systems without Outlook COM support\n\n" +
                "Please install or switch to Outlook Win32 Desktop and try again.",
                "View Inferred Contacts \u2014 COM Required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        Application.Run(new MainForm());
    }
}
