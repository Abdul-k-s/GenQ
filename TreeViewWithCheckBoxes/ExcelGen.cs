using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using static System.String;
using Rectangle = System.Drawing.Rectangle;

namespace GenQ
{
    public class ExcelGen : INotifyPropertyChanged
    {
        #region Fields and Properties
        
        public string filePath = null;
        public ExcelPackage mainPack = null;
        public Document doc = null;
        public string ClientLogo { get; set; }
        public string CreatorLogo { get; set; }

        private string _ProjectName;
        public string ProjectName
        {
            get { return _ProjectName; }
            set
            {
                _ProjectName = value;
                this.OnPropertyChanged("ProjectName");
            }
        }

        private string _ProjectLocation;
        public string ProjectLocation
        {
            get { return _ProjectLocation; }
            set
            {
                _ProjectLocation = value;
                this.OnPropertyChanged("ProjectLocation");
            }
        }

        #endregion

        #region Excel Generation

        public void PrintExcel(List<FooViewModel> ViewTypes)
        {
            List<string> temp = new List<string>();
            List<string> listDiv = new List<string>();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            ExcelPackage BOQexcel = new ExcelPackage();
            ProjectInfo pInfo = doc.ProjectInformation;

            #region Cover Page

            var workSheetIsmail = BOQexcel.Workbook.Worksheets.Add($"Summary");
            workSheetIsmail.Cells.Style.Font.Name = "Calibri";
            workSheetIsmail.Rows.Height = 20;
            workSheetIsmail.Columns.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            workSheetIsmail.Cells["B2:C2,B3:C3,B4:C4,D2:E2,D3:E3,D4:E4,G2:H2,G3:H3,G4:H4,B7:C7,B8:B9,C8:H9,I8:I9"].Merge = true;
            workSheetIsmail.Cells["B2,B3,B4,G2,G3,G4,B7,B8,C8,I8"].Style.Font.Bold = true;

            workSheetIsmail.Cells["B2"].Value = "Project Name:";
            workSheetIsmail.Cells["B3"].Value = "Project Number:";
            workSheetIsmail.Cells["B4"].Value = "Project Address:";
            workSheetIsmail.Cells["G2"].Value = "Client:";
            workSheetIsmail.Cells["G3"].Value = "Consultant:";
            workSheetIsmail.Cells["G4"].Value = "Start Date:";
            workSheetIsmail.Cells["B7"].Value = "Bill Summary:";

            if (ProjectName != null) { workSheetIsmail.Cells["D2"].Value = ProjectName; }
            if (ProjectLocation != null) { workSheetIsmail.Cells["D4"].Value = ProjectLocation; }
            workSheetIsmail.Cells["D3"].Value = pInfo.Number;
            workSheetIsmail.Cells["I2"].Value = pInfo.ClientName;
            workSheetIsmail.Cells["I3"].Value = pInfo.OrganizationName;
            workSheetIsmail.Cells["I4"].Value = pInfo.IssueDate;
            workSheetIsmail.Cells["B8"].Value = "Div No.";
            workSheetIsmail.Cells["C8"].Value = "                                  Division Name";
            workSheetIsmail.Cells["I8"].Value = "      Total Cost";

            workSheetIsmail.Cells["B2:I4,B7:C7"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            workSheetIsmail.Cells["B8:I9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            workSheetIsmail.Cells["J4"].Value = "";

            #endregion

            List<string> Divisions = new List<string>();
            List<ExcelWorksheet> Worksheets = new List<ExcelWorksheet>();

            foreach (FooViewModel type in ViewTypes)
            {
                Divisions.Add($"Div {type.Division.Section.ToString().Substring(0, 2)}");
            }
            Divisions = Divisions.Distinct().ToList();
            Divisions.Sort();
            foreach (string Division in Divisions)
            {
                Worksheets.Add(BOQexcel.Workbook.Worksheets.Add(Division));
            }

            #region SheetsFormating
            foreach (var sheet in Worksheets)
            {
                sheet.TabColor = System.Drawing.Color.Black;
                sheet.DefaultRowHeight = 30;
                sheet.Cells["A1:A3,B1:B2,C1:C3,D1:D3,E1:E3,F1:F3"].Merge = true;

                sheet.Columns.Width = 20;
                sheet.Columns.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Column(2).Width = 50;
                sheet.Column(2).Style.WrapText = true;

                sheet.Row(1).Height = 40;
                sheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(1).Style.Font.Bold = true;
                sheet.Row(1).Style.Font.Size = 14;
                sheet.Cells.Style.Font.Name = "Arial";
                sheet.Cells[3, 2].Style.Font.Size = 10;
                sheet.Cells[3, 2].Style.Font.Bold = true;
                sheet.Cells[3, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A1:B3,A1:F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A1:B3,A1:F1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                sheet.PrinterSettings.PaperSize = ePaperSize.A4;
                sheet.View.PageBreakView = true;
                sheet.PrinterSettings.FitToPage = true;
            }
            #endregion

            #region PrintSheetData
            ViewTypes = ViewTypes.OrderBy(o => o.Section.Section).ToList();
            int rowIndex = 4;
            int e = 0;
            foreach (var worksheet in Worksheets)
            {
                int secindex = 0;
                int subsecindex = 0;
                foreach (FooViewModel Type in ViewTypes)
                {
                    if (Type.Division.Section.ToString().Substring(0, 2) == worksheet.Name.Substring(4, 2))
                    {
                        worksheet.Cells[3, 2].Value = $"{Type.Division.Section} {Type.Division.Name}";
                        worksheet.Cells[3, 2].Style.Font.Size = 12;

                        if (e == 0 || ViewTypes[e].Section.Section != ViewTypes[e - 1].Section.Section)
                        {
                            subsecindex = 0;
                            secindex++;
                            subsecindex++;
                            worksheet.Cells[rowIndex, 1].Value = (secindex);
                            worksheet.Cells[rowIndex, 2].Value = $"{Type.Section.Section} {Type.Section.Name}";
                            worksheet.Cells[rowIndex, 2].Style.Font.UnderLine = true;
                            worksheet.Cells[rowIndex, 2].Style.Font.Bold = true;
                            rowIndex++;

                            worksheet.Cells[rowIndex, 1].Value = $"{secindex}.{subsecindex}";

                            if (Type.Description != null)
                            {
                                worksheet.Cells[rowIndex, 2].Value = $"{Type.Description}\n{Type.Name}";
                            }
                            else { worksheet.Cells[rowIndex, 2].Value = Type.Name; }
                            worksheet.Cells[rowIndex, 3].Value = UnitConverter(Type.UnitType);
                            worksheet.Cells[rowIndex, 4].Value = Type.Quantity;
                            worksheet.Cells[rowIndex, 5].Value = Type.Cost;
                            worksheet.Cells[rowIndex, 6].Formula = $"{worksheet.Cells[rowIndex, 4].Address}*{worksheet.Cells[rowIndex, 5].Address}";
                            rowIndex++;
                            e++;
                        }
                        else
                        {
                            subsecindex++;
                            worksheet.Cells[rowIndex, 1].Value = $"{secindex}.{subsecindex}";

                            if (Type.Description != null)
                            {
                                worksheet.Cells[rowIndex, 2].Value = $"{Type.Description}\n{Type.Name}";
                            }
                            else { worksheet.Cells[rowIndex, 2].Value = Type.Name; }
                            worksheet.Cells[rowIndex, 3].Value = UnitConverter(Type.UnitType);
                            worksheet.Cells[rowIndex, 4].Value = Type.Quantity;
                            worksheet.Cells[rowIndex, 5].Value = Type.Cost;
                            worksheet.Cells[rowIndex, 6].Formula = $"{worksheet.Cells[rowIndex, 4].Address}*{worksheet.Cells[rowIndex, 5].Address}";

                            rowIndex++;
                            e++;
                        }
                    }
                }

                var finalcell = worksheet.Cells[worksheet.Dimension.End.Row + 1, worksheet.Dimension.End.Column];
                finalcell.Formula = $"SUM(F4:F{worksheet.Dimension.End.Row})";
                finalcell.Style.Font.Bold = true;

                var totlcellsheet = worksheet.Cells[worksheet.Dimension.End.Row, worksheet.Dimension.Start.Column];
                totlcellsheet.Value = "Total";
                totlcellsheet.Style.Font.Bold = true;

                worksheet.Cells[$"{totlcellsheet.LocalAddress}:{finalcell.Address}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[$"{totlcellsheet.LocalAddress}:{finalcell.Address}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                temp.Add(worksheet.Cells[worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].FullAddressAbsolute);
                listDiv.Add(worksheet.Cells[3, 2].FullAddressAbsolute);
                rowIndex = 4;
                worksheet.Column(2).BestFit = true;
                worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:3"];

                worksheet.View.PageBreakView = true;
                worksheet.PrinterSettings.FitToPage = true;
                worksheet.PrinterSettings.FitToWidth = 1;
                worksheet.PrinterSettings.FitToHeight = 2;
            }
            #endregion

            #region PrintSheetHeadings
            foreach (var sheet in Worksheets)
            {
                sheet.Cells[1, 1].Value = "Item No.";
                sheet.Cells[1, 2].Value = $"Description";
                sheet.Cells[1, 3].Value = "Unit";
                sheet.Cells[1, 4].Value = "Quantity";
                sheet.Cells[1, 5].Value = "Rate";
                sheet.Cells[1, 6].Value = "Total (EGP)";
            }
            #endregion

            #region CoverPageTotalFillDynamic
            List<string> divisionPlace = new List<string>();
            List<string> totalDivPlace = new List<string>();
            List<string> totalDivNo = new List<string>();
            List<string> MergedH = new List<string>();

            for (int i = 10; i <= 55; i++) 
            { 
                MergedH.Add(($"H{i}")); 
                totalDivPlace.Add($"I{i}"); 
                divisionPlace.Add($"C{i}"); 
                totalDivNo.Add($"B{i}"); 
            }
            int j = 0;

            foreach (string cell in temp)
            {
                workSheetIsmail.Cells[$"{divisionPlace[j]}:{MergedH[j]}"].Merge = true;
                workSheetIsmail.Cells[totalDivNo[j]].Formula = "=\"Div \" " + $"& MID({listDiv[j]},1,2)";
                workSheetIsmail.Cells[totalDivPlace[j]].Formula = $"={temp[j]}";
                workSheetIsmail.Cells[divisionPlace[j]].Formula = $"=REPLACE({listDiv[j]}, 1, 9, )";
                j++;
            }

            var cellTotal = workSheetIsmail.Cells[(workSheetIsmail.Dimension.End.Row + 2), (workSheetIsmail.Dimension.Start.Column) + 1];
            cellTotal.Value = "Total";
            cellTotal.Style.Font.Bold = true;
            var celltotalEq = workSheetIsmail.Cells[((workSheetIsmail.Dimension.End.Row)), (workSheetIsmail.Dimension.End.Column) - 1];
            celltotalEq.Formula = $"SUM(I10:I{j + 9})";
            celltotalEq.Style.Font.Bold = true;
            workSheetIsmail.Cells[$"{cellTotal.LocalAddress}:{celltotalEq.Address}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheetIsmail.Cells[$"{cellTotal.LocalAddress}:{celltotalEq.Address}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            workSheetIsmail.Cells[$"B8:{cellTotal}"].Style.Border.BorderAround(ExcelBorderStyle.Hair);
            workSheetIsmail.Cells[$"I8:{celltotalEq}"].Style.Border.BorderAround(ExcelBorderStyle.Hair);
            workSheetIsmail.Cells[$"{cellTotal}:{celltotalEq}"].Style.Border.BorderAround(ExcelBorderStyle.Medium);

            workSheetIsmail.Cells[(workSheetIsmail.Dimension.End.Row + 1), (workSheetIsmail.Dimension.Start.Column) + 1].Value = "";

            workSheetIsmail.Cells[$"B8:{workSheetIsmail.Cells[(workSheetIsmail.Dimension.End.Row) - 1, (workSheetIsmail.Dimension.End.Column) - 1].LocalAddress}"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            workSheetIsmail.Columns.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            workSheetIsmail.Cells[$"I10:{celltotalEq.Address}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheetIsmail.Cells[(workSheetIsmail.Dimension.End.Row) - 1, (workSheetIsmail.Dimension.End.Column) - 1].Style.Font.Size = 13;
            workSheetIsmail.Cells.Style.Font.Size = 13;
            workSheetIsmail.Column(9).Width = 17;

            workSheetIsmail.View.PageBreakView = true;
            workSheetIsmail.PrinterSettings.FitToPage = true;
            workSheetIsmail.PrinterSettings.FitToWidth = 1;

            #endregion

            #region FileInfo
            if (ProjectName != null) { BOQexcel.Workbook.Properties.Title = $"BOQ of  {ProjectName} "; }
            BOQexcel.Workbook.Properties.Subject = "Bill of quantity";
            BOQexcel.Workbook.Properties.Comments = "Created by GenQ, Revit BOQ plugin";
            BOQexcel.Workbook.Properties.Status = "Ready";
            BOQexcel.Workbook.Properties.Category = "Quantity Surveying";
            #endregion

            #region Image
            workSheetIsmail.InsertRow(1, 3);
            if (ClientLogo != null)
            {
                AddImageToWorksheet(ClientLogo, workSheetIsmail, 0, 5, 1, 0);
            }
            if (CreatorLogo != null)
            {
                AddImageToWorksheet(CreatorLogo, workSheetIsmail, 0, 5, workSheetIsmail.Dimension.End.Column - 2, 20);
            }
            #endregion

            mainPack = BOQexcel;
        }

        #endregion

        #region Image Handling (EPPlus 7.x compatible)

        /// <summary>
        /// Adds an image to the worksheet at specified position (EPPlus 7.x compatible)
        /// </summary>
        private void AddImageToWorksheet(string imgPath, ExcelWorksheet ws, int row, int rowOffset, int col, int colOffset)
        {
            if (string.IsNullOrEmpty(imgPath) || !File.Exists(imgPath))
                return;

            try
            {
                // Generate unique name for the picture
                string pictureName = $"img_{Guid.NewGuid():N}";

                // EPPlus 7.x: Use FileInfo directly
                FileInfo imageFile = new FileInfo(imgPath);
                
                // Add picture using the new EPPlus 7.x API
                ExcelPicture picture = ws.Drawings.AddPicture(pictureName, imageFile);
                
                // Position the image
                picture.SetPosition(row, rowOffset, col, colOffset);

                // Resize to fit header area (approximately 67 pixels height)
                using (Image img = Image.FromFile(imgPath))
                {
                    int targetHeight = 67;
                    if (img.Height > targetHeight)
                    {
                        int percentage = (int)Math.Round((double)targetHeight / img.Height * 100);
                        picture.SetSize(percentage);
                    }
                }

                picture.LockAspectRatio = true;
            }
            catch (Exception ex)
            {
                // Log error but don't fail the export
                System.Diagnostics.Debug.WriteLine($"Failed to add image: {ex.Message}");
            }
        }

        /// <summary>
        /// Resizes an image for header/footer use
        /// </summary>
        private Image ResizeImageForHeader(string filepath)
        {
            if (string.IsNullOrEmpty(filepath) || !File.Exists(filepath))
                return null;

            try
            {
                Image orgImage = Image.FromFile(filepath);
                float fixedHeight = 63f;
                int destHeight, destWidth;

                if (orgImage.Height > fixedHeight)
                {
                    destHeight = (int)fixedHeight;
                    destWidth = (int)(fixedHeight / orgImage.Height * orgImage.Width);
                }
                else
                {
                    destHeight = orgImage.Height;
                    destWidth = orgImage.Width;
                }

                Bitmap bmp = new Bitmap(destWidth, destHeight);
                bmp.SetResolution(orgImage.HorizontalResolution, orgImage.VerticalResolution);
                
                using (Graphics grPhoto = Graphics.FromImage(bmp))
                {
                    grPhoto.DrawImage(orgImage,
                        new Rectangle(0, 0, destWidth, destHeight),
                        new Rectangle(0, 0, orgImage.Width, orgImage.Height),
                        GraphicsUnit.Pixel);
                }

                return bmp;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Save and Export

        /// <summary>
        /// Saves the Excel file and opens it
        /// </summary>
        public void SavenOpen()
        {
            try
            {
                // Save to Documents folder to avoid permission issues with Program Files
                string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string fileName = $"BOQ {ProjectName} {DateTime.Now:ddMMyyyy HHmm}.xlsx";
                string path = Path.Combine(documentsFolder, fileName);

                using (FileStream fStream = File.Create(path))
                {
                    mainPack.SaveAs(fStream);
                }

                var psi = new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true };
                System.Diagnostics.Process.Start(psi);
            }
            catch (IOException)
            {
                Autodesk.Revit.UI.TaskDialog.Show("File opening error", 
                    "There is a file with the same name is being used, try to close it firstly.", 
                    TaskDialogCommonButtons.Retry);
            }
        }

        /// <summary>
        /// Saves as Excel file only (PDF export removed - use Excel's built-in PDF export)
        /// Note: Microsoft.Office.Interop.Excel is not available in .NET 8.
        /// Users should open the Excel file and use File > Save As > PDF.
        /// </summary>
        public void SaveAsExcel(string outputPath = null)
        {
            try
            {
                // Format worksheets for printing/PDF
                for (int i = 1; i < mainPack.Workbook.Worksheets.Count; i++)
                {
                    mainPack.Workbook.Worksheets[i].InsertRow(1, 1);
                    mainPack.Workbook.Worksheets[i].Cells[mainPack.Workbook.Worksheets[i].Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    mainPack.Workbook.Worksheets[i].Cells[mainPack.Workbook.Worksheets[i].Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    mainPack.Workbook.Worksheets[i].Cells[mainPack.Workbook.Worksheets[i].Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    mainPack.Workbook.Worksheets[i].Cells[mainPack.Workbook.Worksheets[i].Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    mainPack.Workbook.Worksheets[i].Cells["E:E,F:F"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                #region CoverPDF
                var coverPDF = mainPack.Workbook.Worksheets.Add("Cover Page");
                coverPDF.Cells["A21:J23,A24:J25,D39:G40"].Merge = true;
                coverPDF.Cells["A21"].Value = ProjectName;
                coverPDF.Cells["A24"].Value = ProjectLocation;
                coverPDF.Cells["D39"].Value = "Bill of Quantities";
                coverPDF.Cells["A21,A24,D39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                coverPDF.Cells["A21,A24,D39"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                coverPDF.Cells["A21,A24,D39"].Style.Font.Bold = true;
                coverPDF.Cells["A21,A24"].Style.Font.Size = 20;
                coverPDF.Cells["D39"].Style.Font.Size = 16;
                coverPDF.Cells["A2:J47"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                coverPDF.Columns[8, 9].Width = 11.3215;

                mainPack.Workbook.Worksheets.MoveToStart("Cover Page");
                #endregion

                int Count = mainPack.Workbook.Worksheets.Count;
                for (int i = 1; i < Count - 1; i++)
                {
                    mainPack.Workbook.Worksheets.Add($"{i}");
                    mainPack.Workbook.Worksheets.MoveBefore(mainPack.Workbook.Worksheets.Count - 1, i + i);
                    mainPack.Workbook.Worksheets[i + i].Cells["A21:J23,A24:J25"].Merge = true;
                    mainPack.Workbook.Worksheets[i + i].Cells["A21"].Formula = "=\"Division \" " + $"& MID({mainPack.Workbook.Worksheets[i + i + 1].Cells["B4"].FullAddressAbsolute},1,2)";
                    mainPack.Workbook.Worksheets[i + i].Cells["A24"].Formula = $"=REPLACE({mainPack.Workbook.Worksheets[i + i + 1].Cells["B4"].FullAddressAbsolute}, 1, 9, )";
                    mainPack.Workbook.Worksheets[i + i].Cells["A21,A24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    mainPack.Workbook.Worksheets[i + i].Cells["A21,A24"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    mainPack.Workbook.Worksheets[i + i].Cells["A21,A24"].Style.Font.Bold = true;
                    mainPack.Workbook.Worksheets[i + i].Cells["A21,A24"].Style.Font.Size = 16;
                    mainPack.Workbook.Worksheets[i + i].Cells["A2:J47"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    mainPack.Workbook.Worksheets[i + i].Columns[8, 9].Width = 11.3215;
                }

                // Configure print settings for all worksheets
                foreach (ExcelWorksheet ws in mainPack.Workbook.Worksheets)
                {
                    ws.PrinterSettings.PaperSize = ePaperSize.A4;
                    ws.HeaderFooter.OddHeader.CenteredText = $"\n{ProjectName}\n{ProjectLocation}";
                    ws.HeaderFooter.EvenHeader.CenteredText = $"\n{ProjectName}\n{ProjectLocation}";
                    ws.HeaderFooter.EvenFooter.RightAlignedText = DateTime.Now.ToString("dd MMM yyyy");
                    ws.HeaderFooter.OddFooter.RightAlignedText = DateTime.Now.ToString("dd MMM yyyy");
                    ws.HeaderFooter.ScaleWithDocument = false;
                    ws.HeaderFooter.AlignWithMargins = false;
                    ws.PrinterSettings.FitToPage = true;
                    ws.PrinterSettings.FitToWidth = 1;
                    ws.PrinterSettings.FitToHeight = 2;
                }

                // Save the file
                string path = outputPath;
                if (string.IsNullOrEmpty(path))
                {
                    // Save to Documents folder to avoid permission issues with Program Files
                    string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string fileName = $"BOQ {ProjectName} {DateTime.Now:ddMMyyyy HHmm}.xlsx";
                    path = Path.Combine(documentsFolder, fileName);
                }

                filePath = path;

                using (FileStream fStream = File.Create(path))
                {
                    mainPack.SaveAs(fStream);
                }
            }
            catch (Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", 
                    $"Failed to save Excel file: {ex.Message}", 
                    TaskDialogCommonButtons.Ok);
            }
        }

        #endregion

        #region Helper Methods

        public string UnitConverter(UnitTypeEnum unit)
        {
            switch (unit)
            {
                case UnitTypeEnum.None:
                    return "None";
                case UnitTypeEnum.Volume:
                    return "m\xB3";
                case UnitTypeEnum.Area:
                    return "m\xB2";
                case UnitTypeEnum.Perimeter:
                    return "m";
                case UnitTypeEnum.Length:
                    return "m";
                case UnitTypeEnum.Count:
                    return "item";
                default:
                    return "none";
            }
        }

        #endregion

        #region INotifyPropertyChanged Members

        void OnPropertyChanged(string prop)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
}
