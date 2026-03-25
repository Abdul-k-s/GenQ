using System.Windows;
using System.Windows.Media;
using GenQ.Data;

namespace GenQ.Views
{
    public partial class SqlConnectionDialog : Window
    {
        public ConnectionSettings Settings { get; private set; }
        public bool IsSaved { get; private set; }

        public SqlConnectionDialog()
        {
            InitializeComponent();
            LoadSettings();
        }

        public SqlConnectionDialog(Window owner) : this()
        {
            Owner = owner;
        }

        private void LoadSettings()
        {
            Settings = ConnectionSettings.Load();
            txtServer.Text = Settings.Server;
            txtDatabase.Text = Settings.Database;
            rbIntegrated.IsChecked = Settings.IntegratedSecurity;
            rbSqlAuth.IsChecked = !Settings.IntegratedSecurity;
            txtUsername.Text = Settings.Username ?? "";
            pnlCredentials.IsEnabled = !Settings.IntegratedSecurity;
        }

        private void AuthenticationChanged(object sender, RoutedEventArgs e)
        {
            if (pnlCredentials != null)
            {
                pnlCredentials.IsEnabled = rbSqlAuth.IsChecked == true;
            }
        }

        private void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            UpdateSettingsFromUI();
            
            txtStatus.Text = "Testing connection...";
            txtStatus.Foreground = Brushes.Gray;
            
            // Force UI update
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Render, 
                new System.Action(() => { }));

            if (Settings.TestConnection(out string errorMessage))
            {
                txtStatus.Text = "Connection successful!";
                txtStatus.Foreground = Brushes.Green;
            }
            else
            {
                txtStatus.Text = $"Connection failed: {errorMessage}";
                txtStatus.Foreground = Brushes.Red;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            UpdateSettingsFromUI();
            Settings.Save();
            IsSaved = true;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void UpdateSettingsFromUI()
        {
            Settings.Server = txtServer.Text.Trim();
            Settings.Database = txtDatabase.Text.Trim();
            Settings.IntegratedSecurity = rbIntegrated.IsChecked == true;
            Settings.Username = txtUsername.Text.Trim();
            Settings.Password = txtPassword.Password;
        }
    }
}
