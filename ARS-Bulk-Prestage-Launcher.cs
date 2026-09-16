using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal static class Program
{
    private const string ScriptName = "ARS-Bulk-Prestage-ADComputers.ps1";

    [STAThread]
    private static void Main()
    {
        string launcherPath = Assembly.GetExecutingAssembly().Location;
        string launcherDirectory = Path.GetDirectoryName(launcherPath);
        string scriptPath = Path.Combine(launcherDirectory, ScriptName);
        string powershellPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.System),
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");

        if (!File.Exists(scriptPath))
        {
            ShowError(
                "The PowerShell script was not found.\r\n\r\n" +
                "Place this launcher beside:\r\n" + ScriptName +
                "\r\n\r\nExpected location:\r\n" + scriptPath);
            return;
        }

        if (!File.Exists(powershellPath))
        {
            ShowError("Windows PowerShell was not found at:\r\n\r\n" + powershellPath);
            return;
        }

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = powershellPath,
                Arguments =
                    "-NoLogo -NoProfile -ExecutionPolicy Bypass -Command " +
                    QuoteArgument(
                        "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force; " +
                        "& '" + scriptPath.Replace("'", "''") + "'"),
                WorkingDirectory = launcherDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            Process.Start(startInfo);
        }
        catch (Exception exception)
        {
            ShowError(
                "The PowerShell script could not be started.\r\n\r\n" +
                exception.Message);
        }
    }

    private static string QuoteArgument(string value)
    {
        var result = new StringBuilder("\"");
        int backslashes = 0;

        foreach (char character in value)
        {
            if (character == '\\')
            {
                backslashes++;
                continue;
            }

            if (character == '"')
            {
                result.Append('\\', (backslashes * 2) + 1);
                result.Append('"');
                backslashes = 0;
                continue;
            }

            result.Append('\\', backslashes);
            backslashes = 0;
            result.Append(character);
        }

        result.Append('\\', backslashes * 2);
        result.Append('"');
        return result.ToString();
    }

    private static void ShowError(string message)
    {
        MessageBox.Show(
            message,
            "ARS Bulk Computer Prestage Launcher",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
    }
}
