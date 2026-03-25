using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using GenQ.Data;
using GenQ.Views;




namespace GenQ
{
    public partial class Window2 : Window
    {

        public Document doc;
        public List<FooViewModel> Tree = new List<FooViewModel>();
        public FooViewModel viewModel = new FooViewModel();
        public List<FooViewModel> CheckedTypes = new List<FooViewModel>();
        public List<FooViewModel> Selected = new List<FooViewModel>();
        public List<LevelViewModel> lvls = new List<LevelViewModel>();

        private List<ReadMasterFormat> _Master; // field
        public List<ReadMasterFormat> Master   // property
        {
            get { return _Master; }   // get method
            set { _Master = value; }  // set method
        }
        public Window2(Document doc)
        {
            this.doc = doc;
            FooViewModel viewModel = new FooViewModel();

            InitializeComponent();


            ReadMasterFormat read = new ReadMasterFormat();
            Master = read.GetfromExcel();
            Master = Master[0].Children;
            LevelViewModel Levels = new LevelViewModel();
            lvls = Levels.Getfromrevit(doc);
            lvls.AddRange(lvls[0].Children);
            //lvls.Add("All Levels");
            //lvls.AddRange(viewModel.GetLevelsNames(doc));
            //this.Levels.SelectedValue = "All Levels";



            Tree = viewModel.Getfromrevit(doc, (bool)IncludeLinks.IsChecked, SelectedLevels(lvls));
            this.tree.ItemsSource = Tree;
            this.Levels.ItemsSource = lvls;
            FooViewModel root = this.tree.Items[0] as FooViewModel;

            this.tree.Focus();



        }

        public List<string> SelectedLevels(List<LevelViewModel> lvls)
        {
            List<string> LevelNames = new List<string>();
            if (lvls[0].IsChecked == true)
            {
                LevelNames.Add(lvls[0].Name);
            }
            else
            {
                foreach (LevelViewModel lvl in lvls[0].Children)
                {
                    if (lvl.IsChecked == true)
                    {
                        LevelNames.Add(lvl.Name);
                    }
                }
            }
            return LevelNames;
        }
        public void OK_Button(object sender, RoutedEventArgs e)
        {
            Selected = GetSelected();
            
            // Show export buttons if selection is valid
            if (Selected.Count > 0)
            {
                btnExport.Visibility = System.Windows.Visibility.Visible;
                btnSqlExport.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {
                btnExport.Visibility = System.Windows.Visibility.Collapsed;
                btnSqlExport.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        public List<FooViewModel> GetSelected()
        {
            List<FooViewModel> exporttypes = new List<FooViewModel>();
            for (int i = Tree.Count - 1; i >= 0; i--)
            {
                if (Tree[i].IsChecked == true)
                {
                    for (int q = CheckedTypes.Count - 1; q >= 0; q--)
                    {
                        if (Tree[i].Name == CheckedTypes[q].Name)
                        {
                            Tree[i].UnitType = CheckedTypes[q].UnitType;
                            Tree[i].Division = CheckedTypes[q].Division;
                            Tree[i].Cost = CheckedTypes[q].Cost;
                            Tree[i].Description = CheckedTypes[q].Description;
                            Tree[i].Section = CheckedTypes[q].Section;
                        }
                    }
                }
                for (int c = Tree[i].Children.Count - 1; c >= 0; c--)
                {
                    if (Tree[i].Children[c].IsChecked == true)
                    {
                        for (int q = CheckedTypes.Count - 1; q >= 0; q--)
                        {
                            for (int w = CheckedTypes[q].Children.Count - 1; w >= 0; w--)
                            {
                                if (Tree[i].Children[c].Name == CheckedTypes[q].Children[w].Name)
                                {
                                    Tree[i].Children[c].UnitType = CheckedTypes[q].Children[w].UnitType;
                                    Tree[i].Children[c].Division = CheckedTypes[q].Children[w].Division;
                                    Tree[i].Children[c].Cost = CheckedTypes[q].Children[w].Cost;
                                    Tree[i].Children[c].Description = CheckedTypes[q].Children[w].Description;
                                    Tree[i].Children[c].Section = CheckedTypes[q].Children[w].Section;
                                }
                            }
                        }
                    }
                    for (int f = Tree[i].Children[c].Children.Count - 1; f >= 0; f--)
                    {
                        if (Tree[i].Children[c].Children[f].IsChecked ==true)
                        {
                            for (int q = CheckedTypes.Count - 1; q >= 0; q--)
                            {
                                for (int w = CheckedTypes[q].Children.Count - 1; w >= 0; w--)
                                {
                                    for (int r = CheckedTypes[q].Children[w].Children.Count - 1; r >= 0; r--)
                                    {
                                        if (Tree[i].Children[c].Children[f].Name == CheckedTypes[q].Children[w].Children[r].Name)
                                        {
                                            Tree[i].Children[c].Children[f].UnitType = CheckedTypes[q].Children[w].Children[r].UnitType;
                                            Tree[i].Children[c].Children[f].Division = CheckedTypes[q].Children[w].Children[r].Division;
                                            Tree[i].Children[c].Children[f].Cost = CheckedTypes[q].Children[w].Children[r].Cost;
                                            Tree[i].Children[c].Children[f].Description = CheckedTypes[q].Children[w].Children[r].Description;
                                            Tree[i].Children[c].Children[f].Section = CheckedTypes[q].Children[w].Children[r].Section;

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            CheckedTypes = viewModel.Getfromrevit(doc, (bool)IncludeLinks.IsChecked, SelectedLevels(lvls));
            for (int i = Tree.Count - 1; i >= 0; i--)
            {
                CheckedTypes[i].UnitType = Tree[i].UnitType;
                CheckedTypes[i].Division = Tree[i].Division;
                CheckedTypes[i].Cost = Tree[i].Cost;
                CheckedTypes[i].Description = Tree[i].Description;
                CheckedTypes[i].Section = Tree[i].Section;

                if (Tree[i].IsChecked == false)
                {
                    CheckedTypes.Remove(CheckedTypes[i]);
                    break;
                }
                for (int c = Tree[i].Children.Count - 1; c >= 0; c--)
                {
                    CheckedTypes[i].Children[c].UnitType = Tree[i].Children[c].UnitType;
                    CheckedTypes[i].Children[c].Division = Tree[i].Children[c].Division;
                    CheckedTypes[i].Children[c].Cost = Tree[i].Children[c].Cost;
                    CheckedTypes[i].Children[c].Description = Tree[i].Children[c].Description;
                    CheckedTypes[i].Children[c].Section = Tree[i].Children[c].Section;
                    for (int f = Tree[i].Children[c].Children.Count - 1; f >= 0; f--)
                    {
                        CheckedTypes[i].Children[c].Children[f].UnitType = Tree[i].Children[c].Children[f].UnitType;
                        CheckedTypes[i].Children[c].Children[f].Division = Tree[i].Children[c].Children[f].Division;
                        CheckedTypes[i].Children[c].Children[f].Cost = Tree[i].Children[c].Children[f].Cost;
                        CheckedTypes[i].Children[c].Children[f].Description = Tree[i].Children[c].Children[f].Description;
                        CheckedTypes[i].Children[c].Children[f].Section = Tree[i].Children[c].Children[f].Section;

                        if (Tree[i].Children[c].Children[f].IsChecked == true)
                        {
                            exporttypes.Add(CheckedTypes[i].Children[c].Children[f]);

                        }
                        if (Tree[i].Children[c].Children[f].IsChecked == false)
                        {
                            CheckedTypes[i].Children[c].Children.Remove(CheckedTypes[i].Children[c].Children[f]);

                        }

                    }
                    if (Tree[i].Children[c].IsChecked == false)
                    {
                        CheckedTypes[i].Children.Remove(CheckedTypes[i].Children[c]);
                    }
                }
            }
            if (CheckedTypes.Count == 0)
            {
                MessageBox.Show("Please Select Types");
                this.checkedtree.ItemsSource = CheckedTypes;
            }
            else
            {
                this.checkedtree.ItemsSource = CheckedTypes[0].Children;
            }

            return exporttypes;
        }

        private static readonly Regex _regex = new Regex("[^0-9.-]+"); //regex that matches disallowed text
        private static bool IsTextAllowed(string text)
        {
            return !_regex.IsMatch(text);
        }
        private new void PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsTextAllowed(e.Text);
        }

        private void TextBoxPasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(String)))
            {
                String text = (String)e.DataObject.GetData(typeof(String));
                if (!IsTextAllowed(text))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        private void IncludeLinks_Click(object sender, RoutedEventArgs e)
        {
            List<FooViewModel> NewTree = new List<FooViewModel>();
            NewTree = viewModel.Getfromrevit(doc, (bool)IncludeLinks.IsChecked, SelectedLevels(lvls));


            for (int i = NewTree.Count - 1; i >= 0; i--)
            {
                for (int q = Tree.Count - 1; q >= 0; q--)
                {
                    if (NewTree[i].Name == Tree[q].Name)
                    {
                        NewTree[i].IsChecked = Tree[q].IsChecked;
                    }
                }

                for (int c = NewTree[i].Children.Count - 1; c >= 0; c--)
                {
                    for (int q = Tree.Count - 1; q >= 0; q--)
                    {
                        for (int w = Tree[q].Children.Count - 1; w >= 0; w--)
                        {
                            if (NewTree[i].Children[c].Name == Tree[q].Children[w].Name)
                            {
                                NewTree[i].Children[c].IsChecked = Tree[q].Children[w].IsChecked;
                            }
                        }
                    }

                    for (int f = NewTree[i].Children[c].Children.Count - 1; f >= 0; f--)
                    {
                        for (int q = Tree.Count - 1; q >= 0; q--)
                        {
                            for (int w = Tree[q].Children.Count - 1; w >= 0; w--)
                            {
                                for (int r = Tree[q].Children[w].Children.Count - 1; r >= 0; r--)
                                {
                                    if (NewTree[i].Children[c].Children[f].Name == Tree[q].Children[w].Children[r].Name)
                                    {
                                        NewTree[i].Children[c].Children[f].IsChecked = Tree[q].Children[w].Children[r].IsChecked;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            Tree = NewTree;
            this.tree.ItemsSource = Tree;
        }

        private void Levels_Selected(object sender, RoutedEventArgs e)
        {
            this.IncludeLinks_Click(sender, e);
        }

        private void CheckAll(object sender, RoutedEventArgs e)
        {
            Tree[0].IsChecked = true;
        }
        private void UnCheckAll(object sender, RoutedEventArgs e)
        {
            Tree[0].IsChecked = false;
        }

        private void checkedtree_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {


            if (e.OriginalSource.ToString() != "System.Windows.Controls.TextBlock")
            {
                e.Handled = true;

                // Create a new event and raise it on the desired UserControl.
                var eventArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
                eventArg.RoutedEvent = UIElement.MouseWheelEvent;
                eventArg.Source = sender;
                var controlToScroll = ScrollContainer;
                controlToScroll.RaiseEvent(eventArg);
            }
        }

        private void To_Revit_Button(object sender, RoutedEventArgs e)
        {
            using (Transaction t = new Transaction(doc, "Save To Revit"))
            {
                t.Start();
                Selected = GetSelected();
                foreach (FooViewModel Type in Selected)
                {
                    Parameter cost = Type.Type.LookupParameter("Cost");

                    cost.Set(Type.Cost);

                    Parameter desc = Type.Type.LookupParameter("Description");

                    desc.Set(Type.Description);
                }
                t.Commit();
                MessageBox.Show("Data Saved To Revit Successfully!!");
            }
        }

        private void From_Revit_Button(object sender, RoutedEventArgs e)
        {
            Selected = GetSelected();
            foreach (FooViewModel Type in Selected)
            {
                Parameter cost = Type.Type.LookupParameter("Cost");
                if (cost.HasValue == true)
                {
                    Type.Cost = double.Parse(cost.AsDouble().ToString());
                }
                Parameter desc = Type.Type.LookupParameter("Description");
                if (desc.HasValue == true)
                {
                    Type.Description = desc.AsString().ToString();
                }
            }
            MessageBox.Show("Data Imported From Revit Successfully!!");
        }

        private void Export_Button(object sender, RoutedEventArgs e)
        {
            Selected = GetSelected();
            List<string> selectedLevels = SelectedLevels(lvls);
            bool allokay = true;
            if(Selected.Count==0)
            {
                allokay = false;
            }
            foreach (FooViewModel T in Selected)
            {
                if (T.Division == null)
                {
                    MessageBox.Show($"Select Division For Type {T.Name}");
                    allokay = false;
                    break;
                }
                if (T.Section == null)
                {
                    MessageBox.Show($"Select Section For Type {T.Name}");
                    allokay = false;
                    break;
                }
                if (T.UnitType == UnitTypeEnum.None)
                {
                    MessageBox.Show($"Select UnitType For Type {T.Name}");
                    allokay = false;
                    break;
                }
            }
            if (allokay)
            {
                BOQ_Generator Bo = new BOQ_Generator();
                List<FooViewModel> Sorted = Bo.Sort(Selected, selectedLevels);
                if (Bo.AllGood == true)
                {
                    PreviewWindow previewWindow = new PreviewWindow(doc, Sorted);
                    previewWindow.ShowDialog();
                }
            }
        }

        private void SqlExport_Button(object sender, RoutedEventArgs e)
        {
            // First validate selection
            Selected = GetSelected();
            List<string> selectedLevels = SelectedLevels(lvls);
            bool allokay = true;
            
            if (Selected.Count == 0)
            {
                MessageBox.Show("Please select types to export.");
                return;
            }
            
            foreach (FooViewModel T in Selected)
            {
                if (T.Division == null)
                {
                    MessageBox.Show($"Select Division For Type {T.Name}");
                    allokay = false;
                    break;
                }
                if (T.Section == null)
                {
                    MessageBox.Show($"Select Section For Type {T.Name}");
                    allokay = false;
                    break;
                }
                if (T.UnitType == UnitTypeEnum.None)
                {
                    MessageBox.Show($"Select UnitType For Type {T.Name}");
                    allokay = false;
                    break;
                }
            }
            
            if (!allokay) return;

            // Show connection settings dialog
            var connectionDialog = new SqlConnectionDialog(this);
            if (connectionDialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                // Process and export to SQL
                BOQ_Generator Bo = new BOQ_Generator();
                List<FooViewModel> Sorted = Bo.Sort(Selected, selectedLevels);
                
                if (!Bo.AllGood) return;

                var settings = ConnectionSettings.Load();
                using (var conn = settings.CreateConnection())
                {
                    var exporter = new BOQDataExporter(conn);
                    
                    // Initialize schema (creates tables if needed)
                    exporter.InitializeSchema();
                    
                    // Create project record
                    string projectName = doc.Title ?? "Untitled Project";
                    string projectNumber = doc.ProjectInformation?.Number ?? "";
                    string filePath = doc.PathName ?? "";
                    
                    int projectId = exporter.CreateProject(projectName, projectNumber, filePath, null);
                    
                    // Export each item
                    foreach (var item in Sorted)
                    {
                        // Extract family and type from the Name (format: "FamilyName : TypeName")
                        string familyName = "";
                        string typeName = item.Name ?? "";
                        if (typeName.Contains(":"))
                        {
                            var parts = typeName.Split(':');
                            familyName = parts[0].Trim();
                            typeName = parts.Length > 1 ? parts[1].Trim() : "";
                        }

                        var boqItem = new BOQItem
                        {
                            ElementId = (int)(item.Type?.Id?.Value ?? 0),
                            Category = item.Type?.Category?.Name ?? "",
                            FamilyName = familyName,
                            TypeName = typeName,
                            Level = string.Join(", ", selectedLevels),
                            CSIDivision = item.Division?.Section ?? "",
                            CSIDescription = item.Section?.Name ?? "",
                            Quantity = (decimal)item.Quantity,
                            Unit = item.UnitType.ToString(),
                            Area = (decimal)item.Area,
                            Volume = (decimal)item.Volume,
                            Length = (decimal)item.Length,
                            Count = (int)item.Count,
                            Remarks = item.Description ?? ""
                        };
                        
                        exporter.InsertItem(projectId, boqItem);
                    }
                    
                    MessageBox.Show($"Successfully exported {Sorted.Count} items to SQL Server!\n\nProject ID: {projectId}", 
                        "SQL Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"SQL Export failed:\n\n{ex.Message}", 
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}