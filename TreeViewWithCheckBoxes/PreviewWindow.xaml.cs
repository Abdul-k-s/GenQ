using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media.Imaging;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;


namespace GenQ
{
    /// <summary>
    /// Interaction logic for PreviewWindow.xaml
    /// Preview and export BOQ to Excel format
    /// Note: PDF export via Interop is not available in .NET 8.
    /// Users can open the Excel file and save as PDF from Excel.
    /// </summary>
    public partial class PreviewWindow : Window
    {
        public Document doc;
        public ExcelGen Ex = new ExcelGen();
        List<FooViewModel> Sorted = new List<FooViewModel>();

        public PreviewWindow(Document doc, List<FooViewModel> Sorted)
        {
            this.doc = doc;
            Ex.doc = doc;
            this.Sorted = Sorted;
            this.DataContext = Ex;
           
            InitializeComponent();
            
            // Initialize web browser to blank
            try
            {
                pdfx.Navigate(new Uri("about:blank"));
            }
            catch
            {
                // WebBrowser control may not be available
            }
        }

        private void ClientLogo_Button(object sender, RoutedEventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Image files|*bmp;*.jpg;*.png";
            op.FilterIndex = 1;
            op.Multiselect = true;

            if (op.ShowDialog() == true)
            {
                imgclient.Source = new BitmapImage(new Uri(op.FileName));
                Ex.ClientLogo = op.FileName;
            }
            
            CleanupTempFile();
        }

        private void CreatorLogo_Button(object sender, RoutedEventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            op.Filter = "Image files|*bmp;*.jpg;*.png";
            op.FilterIndex = 1;
            op.Multiselect = true;

            if (op.ShowDialog() == true)
            {
                imgcreator.Source = new BitmapImage(new Uri(op.FileName));
                Ex.CreatorLogo = op.FileName;
            }

            CleanupTempFile();
        }

        private void ClickPreview(object sender, RoutedEventArgs e)
        {
            // Generate Excel file for preview
            if (Ex.mainPack == null)
            {
                Ex.PrintExcel(Sorted);
                Ex.SaveAsExcel();
            }
            else if (Ex.filePath == null)
            {
                Ex.SaveAsExcel();
            }

            // Open in default viewer
            if (!string.IsNullOrEmpty(Ex.filePath))
            {
                try
                {
                    var psi = new ProcessStartInfo(Ex.filePath) { UseShellExecute = true };
                    Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to open preview: {ex.Message}");
                }
            }
        }

        private void Export_Excel_Button(object sender, RoutedEventArgs e)
        {
            Ex.PrintExcel(Sorted);
            Ex.SavenOpen();
        }

        private void Export_PDF_Button(object sender, RoutedEventArgs e)
        {
            // PDF export via Interop is not available in .NET 8
            // Generate Excel file and prompt user to save as PDF manually
            if (Ex.mainPack == null)
            {
                Ex.PrintExcel(Sorted);
                Ex.SaveAsExcel();
            }
            else if (Ex.filePath == null)
            {
                Ex.SaveAsExcel();
            }

            if (!string.IsNullOrEmpty(Ex.filePath))
            {
                try
                {
                    var psi = new ProcessStartInfo(Ex.filePath) { UseShellExecute = true };
                    Process.Start(psi);
                    
                    MessageBox.Show(
                        "Excel file opened. To save as PDF:\n\n" +
                        "1. Go to File > Save As (or Export)\n" +
                        "2. Select 'PDF' as the file type\n" +
                        "3. Click Save",
                        "Export to PDF",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to open file: {ex.Message}");
                }
            }
        }

        private void TextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            CleanupTempFile();
        }

        /// <summary>
        /// Cleans up temporary files when values change
        /// </summary>
        private void CleanupTempFile()
        {
            if (Ex.filePath != null)
            {
                int retries = 0;
                while (retries < 3)
                {
                    try
                    {
                        System.IO.File.Delete(Ex.filePath);
                        Ex.mainPack = null;
                        Ex.filePath = null;
                        break;
                    }
                    catch (Exception)
                    {
                        retries++;
                        if (retries >= 3)
                        {
                            try
                            {
                                pdfx.Navigate(new Uri("about:blank"));
                            }
                            catch { }
                            
                            MessageBox.Show("Could not update preview. Please close any open files and try again.");
                            break;
                        }
                    }
                }
            }
        }
    }
}
