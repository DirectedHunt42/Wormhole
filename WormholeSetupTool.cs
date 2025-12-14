using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Principal;
using System.Drawing;

namespace WormholeSetup
{
    public class SetupForm : Form
    {
        private TextBox wormholeExePathTextBox;
        private CheckBox forcePandocCheck;
        private CheckBox forceChocoCheck;
        private CheckBox forceFFmpegCheck;
        private CheckBox showDetailsCheck;
        private Button startButton;
        private TextBox statusTextBox;
        private Label versionLabel;
        private ToolTip toolTip;

        public SetupForm()
        {
            this.Text = "Wormhole Setup";
            this.Size = new System.Drawing.Size(400, 420);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;

            toolTip = new ToolTip();

            Label pathLabel = new Label { Text = "Wormhole Exe Path:", Left = 20, Top = 20, Width = 120 };
            this.Controls.Add(pathLabel);

            wormholeExePathTextBox = new TextBox { Text = Path.Combine(Application.StartupPath, "wormhole.exe"), Left = 140, Top = 20, Width = 180 };
            this.Controls.Add(wormholeExePathTextBox);
            toolTip.SetToolTip(wormholeExePathTextBox, "Path to the wormhole.exe file");

            Button browseButton = new Button { Text = "...", Left = 330, Top = 20, Width = 30 };
            browseButton.Click += BrowseForExe;
            this.Controls.Add(browseButton);
            toolTip.SetToolTip(browseButton, "Browse for wormhole.exe");

            Label installSectionLabel = new Label { Text = "Installation Options", Left = 20, Top = 50, Width = 340 };
            this.Controls.Add(installSectionLabel);

            Label sep1 = new Label { AutoSize = false, Height = 2, Left = 20, Top = 70, Width = 340, BorderStyle = BorderStyle.Fixed3D };
            this.Controls.Add(sep1);

            forcePandocCheck = new CheckBox { Text = "Install Pandoc", Left = 20, Top = 80, Width = 200 };
            this.Controls.Add(forcePandocCheck);
            toolTip.SetToolTip(forcePandocCheck, "Force installation of Pandoc even if already installed");

            forceChocoCheck = new CheckBox { Text = "Install Chocolatey", Left = 20, Top = 110, Width = 200 };
            this.Controls.Add(forceChocoCheck);
            toolTip.SetToolTip(forceChocoCheck, "Force installation of Chocolatey even if already installed");

            forceFFmpegCheck = new CheckBox { Text = "Install FFmpeg", Left = 20, Top = 140, Width = 200 };
            this.Controls.Add(forceFFmpegCheck);
            toolTip.SetToolTip(forceFFmpegCheck, "Force installation of FFmpeg even if already installed");

            Label optionsSectionLabel = new Label { Text = "Other Options", Left = 20, Top = 170, Width = 340 };
            this.Controls.Add(optionsSectionLabel);

            Label sep2 = new Label { AutoSize = false, Height = 2, Left = 20, Top = 190, Width = 340, BorderStyle = BorderStyle.Fixed3D };
            this.Controls.Add(sep2);

            showDetailsCheck = new CheckBox { Text = "Show terminal windows", Left = 20, Top = 200, Width = 200 };
            this.Controls.Add(showDetailsCheck);
            toolTip.SetToolTip(showDetailsCheck, "Show terminal windows during installation");

            startButton = new Button { Text = "Start Setup", Left = 20, Top = 230, Width = 100 };
            startButton.Click += async (sender, e) => await RunSetup();
            this.Controls.Add(startButton);
            toolTip.SetToolTip(startButton, "Start the setup process");

            statusTextBox = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Left = 20, Top = 260, Width = 340, Height = 120, ReadOnly = true };
            this.Controls.Add(statusTextBox);

            versionLabel = new Label { Text = "Version 1.0.1", Left = 250, Top = 230, Width = 100 };
            this.Controls.Add(versionLabel);
        }

        private void BrowseForExe(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Executable files (*.exe)|*.exe";
            ofd.Title = "Select wormhole.exe";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                wormholeExePathTextBox.Text = ofd.FileName;
            }
        }

        private async Task RunSetup()
        {
            startButton.Enabled = false;
            statusTextBox.Clear();
            statusTextBox.BackColor = Color.Yellow;
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
            statusTextBox.BackColor = success ? Color.LightGreen : Color.LightCoral;
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
            string exePath = wormholeExePathTextBox.Text;
            if (!File.Exists(exePath))
            {
                UpdateStatus("Specified wormhole.exe not found.");
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