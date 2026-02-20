using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace GenQ
{
    public class ReadMasterFormat
    {
        //bool? _isChecked = false;
        ReadMasterFormat _parent;
        public static List<ReadMasterFormat> Tree { get; set; }
        public List<ReadMasterFormat> GetfromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Assembly assembly = Assembly.GetExecutingAssembly();

            // level1 D
            List<object> divSections = new List<object>();
            List<object> divTitles = new List<object>();
            List<int> divIndex = new List<int>();

            //Level2 B
            List<object> boldSections = new List<object>();
            List<object> boldTitles = new List<object>();
            List<int> boldIndex = new List<int>();
            ExcelPackage excelPackage = new ExcelPackage();

            string[] strr = assembly.Location.Split(new char[] { '\\', '\\' });
            strr[strr.Length - 1] = "CSI-MasterFormat2018.xlsx";
            try
            {
                excelPackage = new ExcelPackage(new MemoryStream
                     (File.ReadAllBytes(String.Join("\\\\", strr))));
            }
            catch (System.IO.FileNotFoundException)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error CSI-MasterFormat", "File Not Found CSI-MasterFormat2018.xlsx Specify file path", TaskDialogCommonButtons.Ok);

                OpenFileDialog MFormatDir = new OpenFileDialog();
                strr[strr.Length - 1] = "";
                MFormatDir.InitialDirectory = String.Join("\\\\", strr);/*Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);*/
                //MFormatDir.DefaultExt = "*.xlsx";
                MFormatDir.Filter = "Excel Files | *.xlsx";
                MFormatDir.ShowDialog();
                try
                {
                    excelPackage = new ExcelPackage(new MemoryStream
                            (File.ReadAllBytes(MFormatDir.FileName)));
                }
                catch (System.ArgumentException)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("Error CSI-MasterFormat", "No Files Were Chosen Specify file path", TaskDialogCommonButtons.Ok);

                    MFormatDir = new OpenFileDialog();
                    MFormatDir.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    //MFormatDir.DefaultExt = "*.xlsx";
                    MFormatDir.Filter = "Excel Files | *.xlsx";
                    MFormatDir.ShowDialog();
                    excelPackage = new ExcelPackage(new MemoryStream
                            (File.ReadAllBytes(MFormatDir.FileName)));

                }
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Error CSI-MasterFormat", "File Not Found CSI-MasterFormat2018.xlsx Specify file path", TaskDialogCommonButtons.Ok);

                OpenFileDialog MFormatDir = new OpenFileDialog();
                MFormatDir.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                //MFormatDir.DefaultExt = "*.xlsx";
                MFormatDir.Filter = "Excel Files | *.xlsx";

                MFormatDir.ShowDialog();
                excelPackage = new ExcelPackage(new MemoryStream
                        (File.ReadAllBytes(MFormatDir.FileName)));
            }

            Stopwatch sw = new Stopwatch();
            foreach (ExcelWorksheet masterFormatWS in excelPackage.Workbook.Worksheets)
            {
                //ExcelWorksheet masterFormatWS = new ExcelWorksheet()/* = excelPackage.Workbook.Worksheets["2018 Numbers and Titles"]*/;  // remove previous foreach and uncomment this to read specific sheets
                //loop all rows
                for (int i = masterFormatWS.Dimension.Start.Row + 6; i <= masterFormatWS.Dimension.End.Row - 1; i++)
                {
                    //loop all columns in a row
                    for (int j = masterFormatWS.Dimension.Start.Column; j <= masterFormatWS.Dimension.End.Column; j++)
                    {
                        //add the cell data to the List
                        if (masterFormatWS.Cells[i, j].Value != null)
                        {
                            //Divisions bolds 14
                            if (masterFormatWS.Cells[i, j].Style.Font.Size == 14 && masterFormatWS.Cells[i, j].Style.Font.Bold == true)
                            {
                                if (j == 2) divTitles.Add(masterFormatWS.Cells[i, j].Value.ToString());
                                if (j == 1) divSections.Add((masterFormatWS.Cells[i, j].Value));

                                divIndex.Add(i - 7);
                                divIndex = divIndex.Distinct().ToList();
                            }

                            //bolds 13
                            if (masterFormatWS.Cells[i, j].Style.Font.Size == 13 && masterFormatWS.Cells[i, j].Style.Font.Bold == true)
                            {
                                if (j == 2) boldTitles.Add(masterFormatWS.Cells[i, j].Value.ToString());
                                if (j == 1) boldSections.Add((masterFormatWS.Cells[i, j].Value));

                                boldIndex.Add(i - 7);
                                boldIndex = boldIndex.Distinct().ToList();
                            }
                        }
                    }
                }
            }

            ReadMasterFormat root = new ReadMasterFormat("All Divisions", "0", -1) { IsInitiallySelected = true };

            for (int i = 0; i < divTitles.Count; i++)
            {
                root.Children.Add(new ReadMasterFormat((divTitles[i].ToString()), (divSections[i].ToString()), divIndex[i]));
            }

            int x = 0;
            foreach (string sub in boldSections)
            {
                foreach (var div in root.Children.Where(div => div.Section.Substring(0, 2) == sub.Substring(0, 2)))
                {
                    div.Children.Add(new ReadMasterFormat((boldTitles[x].ToString()), (boldSections[x].ToString()), boldIndex[x]));
                }

                x++;
            }
            root.Initialize();
            Tree = new List<ReadMasterFormat> { root };
            sw.Stop();
            var xe = sw.Elapsed;
            return Tree;
        }

        ReadMasterFormat(string name, string section, int index)
        {
            this.Name = name;
            this.Children = new List<ReadMasterFormat>();
            this.Section = section;
            this.index = index;
        }

        public ReadMasterFormat()
        {
        }
        void Initialize()
        {
            foreach (ReadMasterFormat child in this.Children)
            {
                child._parent = this;
                child.Initialize();
            }
        }
        public List<ReadMasterFormat> Children { get; set; }

        public bool IsInitiallySelected { get; private set; }

        public string Name { get; private set; }

        private string _Section;
        public string Section
        {
            get => this._Section;
            set => this.SetSection(value, true, true);
        }
        void SetSection(string value, bool updateChildren, bool updateParent)
        {
            if (value == _Section)
                return;

            _Section = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetSection(_Section, true, false));
        }

        private int _index;
        public int index
        {
            get => this._index;
            set => this.SetIndex(value, true, true);
        }
        void SetIndex(int value, bool updateChildren, bool updateParent)
        {
            if (value == _index)
                return;

            _index = value;

            if (updateChildren)
                this.Children.ForEach(c => c.SetIndex(_index, true, false));
        }
    }
}
