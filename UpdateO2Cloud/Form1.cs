using System;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Drawing;
using System.Management;

namespace UpdateO2Cloud
{
    public partial class Form1 : Form
    {
        private MenuStrip menuStrip;
        private ToolStrip toolStrip;
        private Panel contentPanel;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            ApplyOffice2019Theme();
            GetAntivirusName();

            // Empêche le redimensionnement
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // Bordure fixe
            this.MaximizeBox = false;

            // Gestion des arguments passés en ligne de commande
            string[] args = Environment.GetCommandLineArgs();
            if (args.Length > 1)
            {
                HandleCommandLineArguments(args);
            }
            else
            {
                // Si aucun argument n'est passé, ne rien faire
            }
        }

        private void GetAntivirusName()
        {
            try
            {
                string antivirusName = "Aucun antivirus trouvé.";
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM AntivirusProduct");

                foreach (ManagementObject obj in searcher.Get())
                {
                    antivirusName = obj["displayName"]?.ToString() ?? "Inconnu";
                    break; // Prendre le premier antivirus trouvé
                }

                labelAntivirus.Text = $"Antivirus : {antivirusName}";
            }
            catch (Exception ex)
            {
                labelAntivirus.Text = $"Erreur : {ex.Message}";
            }
        }

        private void InitializeCustomComponents()
        {
            // Configuration de la fenêtre principale
            this.Text = "UpdateO2Cloud";

            // Création du menu
            menuStrip = new MenuStrip();
            ToolStripMenuItem fileMenu = new ToolStripMenuItem("Fichier");
            fileMenu.DropDownItems.AddRange(new ToolStripItem[] {
                new ToolStripMenuItem("Quitter", null, (s, e) => Application.Exit())
            });
            menuStrip.Items.Add(fileMenu);
            this.Controls.Add(menuStrip);

            // Création de la barre d'outils
            toolStrip = new ToolStrip();
            toolStrip.Items.AddRange(new ToolStripItem[] {
                new ToolStripButton("Quitter", null, (s, e) => Application.Exit())
            });
            this.Controls.Add(toolStrip);

            // Panel principal pour le contenu
            contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            this.Controls.Add(contentPanel);
        }

        private void ApplyOffice2019Theme()
        {
            // Couleurs du thème Office 2019
            Color primaryColor = Color.FromArgb(255, 255, 255);      // Blanc
            Color accentColor = Color.FromArgb(0, 120, 215);         // Bleu Office
            Color menuBackColor = Color.FromArgb(245, 246, 247);     // Gris clair
            Color borderColor = Color.FromArgb(229, 229, 229);       // Gris bordure

            // Application du thème
            this.BackColor = primaryColor;
            menuStrip.BackColor = menuBackColor;
            menuStrip.ForeColor = Color.Black;
            toolStrip.BackColor = menuBackColor;
            toolStrip.ForeColor = Color.Black;
            contentPanel.BackColor = primaryColor;

            // Style des boutons de la barre d'outils
            foreach (ToolStripItem item in toolStrip.Items)
            {
                if (item is ToolStripButton button)
                {
                    button.BackColor = menuBackColor;
                    button.ForeColor = Color.Black;
                }
            }
        }

        private void HandleCommandLineArguments(string[] args)
        {
            // Vérifie les arguments passés en ligne de commande
            if (args.Contains("/silent"))
            {
                // Mode silencieux - peut-être faire quelque chose en arrière-plan
                // Si les arguments ne sont pas valides
                MessageBox.Show("Arguments en ligne de commande incorrects. Veuillez utiliser des arguments valides :\n" +
                                "/silent\n" +
                                "/officeupdate\n" +
                                "/officever\n" +
                                "/winver\n" +
                                "/winupdate\n" +
                                "/avsecname\n" +
                                "/avsecupdate",
                                "Erreur",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information); Application.Exit();
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/officeupdate"))
            {
                button1_Click(null, null);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/officever"))
            {
                button2_Click(null, null);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/winver"))
            {
                button3_Click(null, null);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/winupdate"))
            {
                button4_Click(null, null);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/avsecupdate"))
            {
                button6_Click(null, null);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
            else if (args.Contains("/avsecname"))
            {
                string antivirusName = GetAntivirusNameSilent();
                MessageBox.Show($"Nom de l'antivirus détecté : {antivirusName}", "Nom de l'antivirus", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0);  // Force la fermeture de l'application
            }

            else
            {
                // Si les arguments ne sont pas valides
                MessageBox.Show("Arguments en ligne de commande incorrects. Veuillez utiliser des arguments valides :\n" +
                                "/silent\n" +
                                "/officeupdate\n" +
                                "/officever\n" +
                                "/winver\n" +
                                "/winupdate\n" +
                                "/avsecname\n" +
                                "/avsecupdate",
                                "Erreur",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                Environment.Exit(0);  // Force la fermeture de l'application
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = @"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe";
                process.StartInfo.Arguments = "/update user"; // Arguments de la commande
                process.StartInfo.UseShellExecute = false;   // Ne pas utiliser l'exécution du shell Windows
                process.StartInfo.RedirectStandardOutput = true; // Capture la sortie standard
                process.StartInfo.RedirectStandardError = true;  // Capture les erreurs
                process.StartInfo.CreateNoWindow = true;         // Pas de fenêtre de console
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                    MessageBox.Show($"Sortie :\n{output}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show($"Erreur :\n{error}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une exception s'est produite :\n{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = @"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\IntegratedOffice.exe";
                process.StartInfo.Arguments = ""; // Arguments de la commande
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                    MessageBox.Show($"Sortie :\n{output}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show($"Erreur :\n{error}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une exception s'est produite :\n{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = "/c control update";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                    MessageBox.Show($"Sortie :\n{output}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show($"Erreur :\n{error}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une exception s'est produite :\n{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = "/c winver";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                    MessageBox.Show($"Sortie :\n{output}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show($"Erreur :\n{error}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une exception s'est produite :\n{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.Arguments = "/c \"C:\\Program Files\\Bitdefender\\Bitdefender Security\\updcenter\\updcenter.exe\"";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                    MessageBox.Show($"Sortie :\n{output}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (!string.IsNullOrEmpty(error))
                    MessageBox.Show($"Erreur :\n{error}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Une exception s'est produite :\n{ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetAntivirusNameSilent()
        {
            try
            {
                string antivirusName = "Aucun antivirus trouvé.";
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\SecurityCenter2", "SELECT * FROM AntivirusProduct");

                foreach (ManagementObject obj in searcher.Get())
                {
                    antivirusName = obj["displayName"]?.ToString() ?? "Inconnu";
                    break; // Prendre le premier antivirus trouvé
                }

                return antivirusName;  // Retourner le nom de l'antivirus
            }
            catch (Exception ex)
            {
                return $"Erreur : {ex.Message}";  // Retourner l'erreur si une exception se produit
            }
        }

    }
}
