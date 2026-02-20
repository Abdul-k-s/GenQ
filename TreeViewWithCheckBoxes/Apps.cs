using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace GenQ
{
    /// <summary>
    /// GenQ Application Entry Point - Implements IExternalApplication for Revit 2025
    /// Creates the ribbon tab and button for the BOQ Generator
    /// </summary>
    internal class Apps : IExternalApplication
    {
        private const string TAB_NAME = "ITI - GenQ";
        private const string PANEL_NAME = "BOQ Generator";
        private const string BUTTON_NAME = "GenQ";
        private const string BUTTON_TEXT = "GenQ";
        private const string COMMAND_CLASS = "GenQ.Class1";

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // Create ribbon tab
                try
                {
                    application.CreateRibbonTab(TAB_NAME);
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException)
                {
                    // Tab already exists, continue
                }

                // Create ribbon panel
                RibbonPanel ribPanel = application.CreateRibbonPanel(TAB_NAME, PANEL_NAME);

                // Create push button data
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                PushButtonData buttonData = new PushButtonData(
                    BUTTON_NAME,
                    BUTTON_TEXT,
                    assemblyPath,
                    COMMAND_CLASS)
                {
                    ToolTip = "Create Quantity Takeoff and Generate BOQ",
                    LongDescription = "Generate Bill of Quantities from Revit model elements. " +
                                      "Select categories, assign CSI MasterFormat divisions, " +
                                      "specify unit types, and export to Excel."
                };

                // Add button to panel
                PushButton button = ribPanel.AddItem(buttonData) as PushButton;

                // Set button icon
                if (button != null)
                {
                    try
                    {
                        Uri uri = new Uri("pack://application:,,,/GenQ;component/Resources/GenQ.png");
                        button.LargeImage = new BitmapImage(uri);
                    }
                    catch (Exception)
                    {
                        // Icon not found - button will work without icon
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("GenQ Startup Error",
                    $"Failed to initialize GenQ add-in.\n\nError: {ex.Message}");
                return Result.Failed;
            }
        }
    }
}
