using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Principal;

namespace WormholeSetup
{
    public class SetupForm : Form
    {
        private CheckBox forcePandocCheck;
        private CheckBox forceChocoCheck;
        private CheckBox forceFFmpegCheck;
        private CheckBox showDetailsCheck;
        private Button startButton;
        private TextBox statusTextBox;

        public SetupForm()
        {
            this.Text = "Wormhole Setup";
            this.Size = new System.Drawing.Size(400, 300);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;

            forcePandocCheck = new CheckBox { Text = "Force reinstall Pandoc", Left = 20, Top = 20, Width = 200 };
            forceChocoCheck = new CheckBox { Text = "Force reinstall Chocolatey", Left = 20, Top = 50, Width = 200 };
            forceFFmpegCheck = new CheckBox { Text = "Force reinstall FFmpeg", Left = 20, Top = 80, Width = 200 };
            showDetailsCheck = new CheckBox { Text = "Show terminal windows", Left = 20, Top = 110, Width = 200 };

            startButton = new Button { Text = "Start Setup", Left = 20, Top = 140, Width = 100 };
            startButton.Click += async (sender, e) => await RunSetup();

            statusTextBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Left = 20, Top = 180, Width = 340, Height = 80, ReadOnly = true };

            this.Controls.Add(forcePandocCheck);
            this.Controls.Add(forceChocoCheck);
            this.Controls.Add(forceFFmpegCheck);
            this.Controls.Add(showDetailsCheck);
            this.Controls.Add(startButton);
            this.Controls.Add(statusTextBox);
        }

        private async Task RunSetup()
        {
            startButton.Enabled = false;
            statusTextBox.Clear();
            bool success = true;

            try
            {
                UpdateStatus("Starting setup...");

                // Install Pandoc
                UpdateStatus("Handling Pandoc...");
                if (!await InstallPandoc(forcePandocCheck.Checked, showDetailsCheck.Checked))
                {
                    success = false;
                }

                // Install Chocolatey if needed
                UpdateStatus("Handling Chocolatey...");
                if (!await InstallChocolatey(forceChocoCheck.Checked, showDetailsCheck.Checked))
                {
                    success = false;
                }

                // Install FFmpeg
                UpdateStatus("Handling FFmpeg...");
                if (!await InstallFFmpeg(forceFFmpegCheck.Checked, showDetailsCheck.Checked))
                {
                    success = false;
                }

                // Register context menu
                if (success)
                {
                    UpdateStatus("Registering context menu...");
                    if (!RegisterWormhole())
                    {
                        success = false;
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error: {ex.Message}");
                success = false;
            }

            UpdateStatus(success ? "Setup completed successfully." : "Setup failed.");
            startButton.Enabled = true;
        }

        private void UpdateStatus(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateStatus(message)));
                return;
            }

            if (!string.IsNullOrEmpty(message))
            {
                statusTextBox.AppendText(message + Environment.NewLine);
            }
            Application.DoEvents();
        }

        private async Task<bool> InstallPandoc(bool force, bool show)
        {
            if (!force && IsProgramInstalled("pandoc", "--version"))
            {
                UpdateStatus("Pandoc already installed.");
                return true;
            }

            UpdateStatus("Downloading latest Pandoc MSI...");
            string msiPath = Path.Combine(Path.GetTempPath(), "pandoc-latest-windows-x86_64.msi");

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "WormholeSetup");
                string releasesUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest";
                string json = await client.GetStringAsync(releasesUrl);
                JsonDocument doc = JsonDocument.Parse(json);
                string downloadUrl = null;

                foreach (var asset in doc.RootElement.GetProperty("assets").EnumerateArray())
                {
                    string name = asset.GetProperty("name").GetString();
                    if (name.EndsWith("-windows-x86_64.msi"))
                    {
                        downloadUrl = asset.GetProperty("browser_download_url").GetString();
                        break;
                    }
                }

                if (downloadUrl == null)
                {
                    UpdateStatus("Failed to find Pandoc MSI.");
                    return false;
                }

                byte[] msiBytes = await client.GetByteArrayAsync(downloadUrl);
                File.WriteAllBytes(msiPath, msiBytes);
            }

            UpdateStatus("Installing Pandoc...");
            string logPath = Path.Combine(Path.GetTempPath(), "pandoc_install.log");
            string args = $"/i \"{msiPath}\" ALLUSERS=1 /norestart /L*V \"{logPath}\"";
            args += show ? " /qb" : " /qn";

            ProcessStartInfo psi = new ProcessStartInfo("msiexec.exe", args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = !show,
                WindowStyle = show ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
            };

            using (Process proc = new Process { StartInfo = psi })
            {
                proc.OutputDataReceived += (s, e) => UpdateStatus(e.Data);
                proc.ErrorDataReceived += (s, e) => UpdateStatus(e.Data);

                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                await Task.Run(() => proc.WaitForExit());

                if (proc.ExitCode != 0)
                {
                    UpdateStatus($"Pandoc install failed. Code: {proc.ExitCode}. Log: {logPath}");
                    return false;
                }
            }

            // Verify
            string[] pandocPaths = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Pandoc", "pandoc.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Pandoc", "pandoc.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Pandoc", "pandoc.exe")
            };

            foreach (string path in pandocPaths)
            {
                if (File.Exists(path) && RunProcess(path, "--version", true) == 0)
                {
                    UpdateStatus("Pandoc installed successfully.");
                    return true;
                }
            }

            UpdateStatus("Pandoc installed but not verifiable.");
            return false;
        }

        private async Task<bool> InstallChocolatey(bool force, bool show)
        {
            if (!force && IsProgramInstalled("choco", "-v"))
            {
                UpdateStatus("Chocolatey already installed.");
                return true;
            }

            UpdateStatus("Installing Chocolatey...");
            string logPath = Path.Combine(Path.GetTempPath(), "choco_install.log");

            string psScript = @"
Set-ExecutionPolicy Bypass -Scope Process -Force;
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'));
";

            int exitCode = await RunPowerShellScript(psScript, show, logPath);
            if (exitCode != 0)
            {
                UpdateStatus($"Chocolatey install failed. Code: {exitCode}. Log: {logPath}");
                return false;
            }

            // Refresh env
            Environment.SetEnvironmentVariable("Path", Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine) + ";" + Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User));

            if (IsProgramInstalled("choco", "-v"))
            {
                UpdateStatus("Chocolatey installed successfully.");
                return true;
            }

            UpdateStatus("Chocolatey installed but not verifiable.");
            return false;
        }

        private async Task<bool> InstallFFmpeg(bool force, bool show)
        {
            if (!force && IsProgramInstalled("ffmpeg", "-version"))
            {
                UpdateStatus("FFmpeg already installed.");
                return true;
            }

            // Ensure Choco
            if (!await InstallChocolatey(forceChocoCheck.Checked, show))
            {
                return false;
            }

            UpdateStatus("Installing FFmpeg...");
            string logPath = Path.Combine(Path.GetTempPath(), "ffmpeg_install.log");
            string chocoPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "chocolatey", "bin", "choco.exe");
            if (!File.Exists(chocoPath))
            {
                chocoPath = "choco.exe"; // fallback to PATH
            }

            string args = "install ffmpeg -y";
            ProcessStartInfo psi = new ProcessStartInfo(chocoPath, args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = !show,
                WindowStyle = show ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
            };

            using (Process proc = new Process { StartInfo = psi })
            {
                proc.OutputDataReceived += (s, e) => UpdateStatus(e.Data);
                proc.ErrorDataReceived += (s, e) => UpdateStatus(e.Data);

                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                await Task.Run(() => proc.WaitForExit());

                if (proc.ExitCode != 0)
                {
                    UpdateStatus($"FFmpeg install failed. Code: {proc.ExitCode}. Log: {logPath}");
                    return false;
                }
            }

            // Refresh env
            Environment.SetEnvironmentVariable("Path", Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.Machine) + ";" + Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User));

            string ffmpegPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "chocolatey", "bin", "ffmpeg.exe");
            if (File.Exists(ffmpegPath) && RunProcess(ffmpegPath, "-version", true) == 0)
            {
                UpdateStatus("FFmpeg installed successfully.");
                return true;
            }

            UpdateStatus("FFmpeg installed but not verifiable.");
            return false;
        }

        private bool RegisterWormhole()
        {
            string exePath = Path.Combine(Application.StartupPath, "wormhole.exe");
            if (!File.Exists(exePath))
            {
                UpdateStatus("wormhole.exe not found in the same directory.");
                return false;
            }

            int exitCode = RunProcess(exePath, "--register", false);
            if (exitCode == 0)
            {
                UpdateStatus("Context menu registered successfully.");
                return true;
            }
            else
            {
                UpdateStatus($"Registration failed. Code: {exitCode}");
                return false;
            }
        }

        private bool IsProgramInstalled(string program, string args)
        {
            return RunProcess(program, args, true) == 0;
        }

        private int RunProcess(string fileName, string args, bool hidden)
        {
            ProcessStartInfo psi = new ProcessStartInfo(fileName, args)
            {
                UseShellExecute = false,
                CreateNoWindow = hidden,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using (Process proc = new Process { StartInfo = psi })
            {
                try
                {
                    proc.Start();
                    proc.WaitForExit();
                    return proc.ExitCode;
                }
                catch
                {
                    return -1;
                }
            }
        }

        private async Task<int> RunPowerShellScript(string script, bool show, string logPath)
        {
            string tempPsFile = Path.Combine(Path.GetTempPath(), "temp.ps1");
            File.WriteAllText(tempPsFile, script);

            string args = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempPsFile}\"";
            ProcessStartInfo psi = new ProcessStartInfo("powershell.exe", args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = !show,
                WindowStyle = show ? ProcessWindowStyle.Normal : ProcessWindowStyle.Hidden
            };

            using (Process proc = new Process { StartInfo = psi })
            {
                proc.OutputDataReceived += (s, e) => UpdateStatus(e.Data);
                proc.ErrorDataReceived += (s, e) => UpdateStatus(e.Data);

                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                await Task.Run(() => proc.WaitForExit());
                File.Delete(tempPsFile);
                return proc.ExitCode;
            }
        }

        [STAThread]
        static void Main()
        {
            // Check if running as administrator
            bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

            if (!isAdmin)
            {
                // Restart the application with admin privileges
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = Application.ExecutablePath,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                try
                {
                    Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to elevate privileges: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                Application.Exit();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm());
        }
    }
}