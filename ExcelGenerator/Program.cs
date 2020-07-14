using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using Cs = DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System.Data;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;
using Position = DocumentFormat.OpenXml.Drawing.Spreadsheet.Position;
using System.Drawing;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using NumberingFormat = DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat;

namespace ExcelGenerator
{
    static class Program
    {
        public static string ImageFile = AppDomain.CurrentDomain.BaseDirectory + "main_logo.png";

        static string[] saColumnName = CalculateExcelColumnName(50);

        const int cnDataWidth = 1;
        const int cnDataHeight = 3;
        static double[,] faChartData = new double[cnDataHeight, cnDataWidth];
        static string[] saCategories = new string[cnDataWidth];
        static string[] saLegend = new string[cnDataHeight];
        const int cnCategoryRowIndex = 2;
        // make sure this is at least 2 because the "legend" is one less
        const int cnCategoryStartColumnIndex = 3;
        // cnCategoryRowIndex + 1
        const int cnDataStartRowIndex = 3;

        // Charts are complicated enough, but charts have two parts:
        // 1) The actual chart content
        // 2) An image of the chart (as part of DrawingsPart)
        // So it acts like an image inserted into the worksheet too.
        // For simplicity, positioning will be hardcoded.
        const int cnFromRowPosition = 7;
        const int cnFromColumnPosition = 1;
        const int cnToRowPosition = 22;
        const int cnToColumnPosition = 9;

        // axis IDs just need to be unique within the chart itself,
        // so any unique set of values will do.
        const uint cnAxisId1 = 1;
        const uint cnAxisId2 = 2;
        const uint cnAxisId3 = 3;

        public static void Main(string[] args)
        {
            WorkbookPart workbookPart = null;

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var excel = SpreadsheetDocument.Create(memoryStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
                    {
                        workbookPart = excel.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        uint sheetId = 1;
                        excel.WorkbookPart.Workbook.Sheets = new Sheets();
                        Sheets sheets = excel.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                        WorkbookStylesPart stylesPart = excel.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                        stylesPart.Stylesheet = GenerateStyleSheet();
                        stylesPart.Stylesheet.Save();

                        int generatorCount = 3;
                        int sheetCount = ++generatorCount;
                        string[] strSheetName = new string[] { "Plant", "2213979", "2213963", "2213969" };


                        for (int i = 0; i < sheetCount; i++)
                        {
                            string relationshipId = "rId" + (i + 1).ToString();
                            WorksheetPart wSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                            string sheetName = strSheetName[i];
                            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                            sheets.Append(sheet);

                            Worksheet worksheet = new Worksheet();

                            wSheetPart.Worksheet = worksheet;

                            SheetData sheetData = new SheetData();
                            worksheet.Append(sheetData);

                            if (sheetId == 1)
                            {
                                Drawing drawing1 = new Drawing() { Id = relationshipId };
                                worksheet.Append(drawing1);
                                DrawingsPart drawingsPart1 = wSheetPart.AddNewPart<DrawingsPart>(relationshipId);
                                InsertImage(wSheetPart, drawingsPart1, 0, 1, 2, 4, new FileStream(ImageFile, FileMode.Open));

                                saCategories[0] = "Apple";
                                saLegend[0] = "North";
                                faChartData[0, 0] = 6883.16115265859;
                                saLegend[1] = "South";
                                faChartData[1, 0] = 4107.25939045998;
                                saLegend[2] = "East";
                                faChartData[2, 0] = 3641.32064703913;


                                AddToCell(sheetData, 8, 2, "F", CellValues.String, "AA Yarn");
                                AddToCell(sheetData, 1, 2, "I", CellValues.String, "16 February, 2020");

                                //create a MergeCells class to hold each MergeCell
                                MergeCells mergeCells = new MergeCells();

                                //append a MergeCell to the mergeCells for each set of merged cells
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B5:C5") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B6:C6") });

                                wSheetPart.Worksheet.InsertAfter(mergeCells, wSheetPart.Worksheet.Elements<SheetData>().First());


                                AddToCell(sheetData, 9, 5, "B", CellValues.String, "Your Industry");
                                AddToCell(sheetData, 9, 5, "C", CellValues.String, "");

                                AddToCell(sheetData, 5, 6, "B", CellValues.String, "Our Energy");

                                AddToCell(sheetData, 8, 8, "F", CellValues.String, "Plant Summary Report");

                                AddToCell(sheetData, 1, 10, "B", CellValues.String, "Number of generators");
                                AddToCell(sheetData, 0, 10, "G", CellValues.Number, "3");


                                AddToCell(sheetData, 1, 11, "B", CellValues.String, "Total running hours");
                                AddToCell(sheetData, 0, 11, "G", CellValues.String, "56 hours");


                                AddToCell(sheetData, 1, 12, "B", CellValues.String, "Total unplanned generator shutdown");
                                AddToCell(sheetData, 0, 12, "G", CellValues.String, "10 hours");

                                BuildChart(drawingsPart1, sheetName, 14, 1, 27, 8);
                                BuildChart(drawingsPart1, sheetName, 14, 10, 27, 18);


                            }

                            if (sheetId != 1)
                            {
                                AddToCell(sheetData, 8, 2, "B", CellValues.String, "Generator " + sheetName);
                                AddToCell(sheetData, 0, 2, "F", CellValues.String, "16 February, 2020");


                                AddToCell(sheetData, 1, 3, "B", CellValues.String, "Model");
                                AddToCell(sheetData, 0, 3, "D", CellValues.String, "CG-170-16");
                                AddToCell(sheetData, 1, 3, "F", CellValues.String, "Running hour");
                                AddToCell(sheetData, 0, 3, "J", CellValues.String, "22 hrs");
                                AddToCell(sheetData, 1, 3, "L", CellValues.String, "Overhauling due in");
                                AddToCell(sheetData, 0, 3, "O", CellValues.String, "1877 hrs");

                                AddToCell(sheetData, 1, 4, "B", CellValues.String, "Rated Kilowatt");
                                AddToCell(sheetData, 0, 4, "D", CellValues.Number, "1500");
                                AddToCell(sheetData, 1, 4, "F", CellValues.String, "Unplanned shutdown hours");
                                AddToCell(sheetData, 0, 4, "J", CellValues.String, "2 hrs");

                                Drawing drawing1 = new Drawing() { Id = relationshipId };
                                worksheet.Append(drawing1);

                                string chartTitle = "Generator Utilisation";

                                Dictionary<string, int> data1 = new Dictionary<string, int>();
                                data1.Add("Running", 92);
                                data1.Add("Stopped", 8);

                                DrawingsPart drawingsPart = wSheetPart.AddNewPart<DrawingsPart>(relationshipId);
                                InsertPieChartInSpreadSheet(drawingsPart, chartTitle, data1, 6, 0, 18, 6);

                                string title = "Power generation (Kilo Watt Hour)";
                                //Dictionary<DateTime, int> data = new Dictionary<DateTime, int>();
                                //data.Add(Convert.ToDateTime("01-Jan-2019"), 999);
                                //data.Add(Convert.ToDateTime("02-Jan-2019"), 983);
                                //data.Add(Convert.ToDateTime("03-Jan-2019"), 945);
                                //data.Add(Convert.ToDateTime("04-Jan-2019"), 975);
                                //data.Add(Convert.ToDateTime("05-Jan-2019"), 950);
                                //data.Add(Convert.ToDateTime("06-Jan-2019"), 964);
                                //data.Add(Convert.ToDateTime("07-Jan-2019"), 989);
                                //data.Add(Convert.ToDateTime("08-Jan-2019"), 973);
                                //data.Add(Convert.ToDateTime("09-Jan-2019"), 954);
                                //data.Add(Convert.ToDateTime("10-Jan-2019"), 957);
                                //data.Add(Convert.ToDateTime("11-Jan-2019"), 905);
                                //data.Add(Convert.ToDateTime("12-Jan-2019"), 946);
                                //data.Add(Convert.ToDateTime("13-Jan-2019"), 998);
                                //data.Add(Convert.ToDateTime("14-Jan-2019"), 937);
                                //data.Add(Convert.ToDateTime("15-Jan-2019"), 945);
                                //data.Add(Convert.ToDateTime("16-Jan-2019"), 975);
                                //data.Add(Convert.ToDateTime("17-Jan-2019"), 950);
                                //data.Add(Convert.ToDateTime("18-Jan-2019"), 932);
                                //data.Add(Convert.ToDateTime("19-Jan-2019"), 947);
                                //data.Add(Convert.ToDateTime("20-Jan-2019"), 921);
                                //data.Add(Convert.ToDateTime("21-Jan-2019"), 984);
                                //data.Add(Convert.ToDateTime("22-Jan-2019"), 932);
                                //data.Add(Convert.ToDateTime("23-Jan-2019"), 945);
                                //data.Add(Convert.ToDateTime("24-Jan-2019"), 946);


                                Dictionary<string, int> data = new Dictionary<string, int>();
                                data.Add("01/01/2019", 999);
                                data.Add("02/01/2019", 983);
                                data.Add("03/01/2019", 945);
                                data.Add("04/01/2019", 975);
                                data.Add("05/01/2019", 950);
                                data.Add("06/01/2019", 964);
                                data.Add("07/01/2019", 989);
                                data.Add("08/01/2019", 973);
                                data.Add("09/01/2019", 954);
                                data.Add("10/01/2019", 957);
                                data.Add("11/01/2019", 905);
                                data.Add("12/01/2019", 946);
                                data.Add("13/01/2019", 998);
                                data.Add("14/01/2019", 937);
                                data.Add("15/01/2019", 945);
                                data.Add("16/01/2019", 975);
                                data.Add("17/01/2019", 950);
                                data.Add("18/01/2019", 932);
                                data.Add("19/01/2019", 947);
                                data.Add("20/01/2019", 921);
                                data.Add("21/01/2019", 984);
                                data.Add("22/01/2019", 932);
                                data.Add("23/01/2019", 945);
                                data.Add("24/01/2019", 946);


                                InsertBarChartInSpreadsheet(worksheet, drawingsPart, title, data, 6, 8, 22, 20);

                                AddToCell(sheetData, 1, 19, "B", CellValues.String, "Running(%)");
                                AddToCell(sheetData, 0, 19, "D", CellValues.Number, "85.54");
                                AddToCell(sheetData, 1, 19, "E", CellValues.String, "Stopped(%)");
                                AddToCell(sheetData, 0, 19, "G", CellValues.Number, "14.46");

                                //create a MergeCells class to hold each MergeCell
                                MergeCells mergeCells = new MergeCells();

                                //append a MergeCell to the mergeCells for each set of merged cells
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B23:I23") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B25:C25") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("D25:G25") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("H25:I25") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B26:C26") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("D26:G26") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("H26:I26") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("B27:C27") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("D27:G27") });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("H27:I27") });

                                wSheetPart.Worksheet.InsertAfter(mergeCells, wSheetPart.Worksheet.Elements<SheetData>().First());

                                AddToCell(sheetData, 0, 22, "K", CellValues.String, "Monthly power generation");
                                AddToCell(sheetData, 7, 23, "B", CellValues.String, "Alerts");
                                AddToCell(sheetData, 0, 24, "K", CellValues.String, "Shutdown history table");
                                AddToCell(sheetData,10, 25, "B", CellValues.String, "TIME");
                                AddToCell(sheetData, 10, 25, "D", CellValues.String, "ALERT");
                                AddToCell(sheetData, 10, 25, "H", CellValues.String, "PRIORITY");
                                AddToCell(sheetData, 0, 25, "K", CellValues.String, "Date");
                                AddToCell(sheetData, 0, 25, "L", CellValues.String, "Shutdown Time");
                                AddToCell(sheetData, 0, 25, "M", CellValues.String, "Duration");
                                AddToCell(sheetData, 0, 25, "N", CellValues.String, "Reason");

                                AddToCell(sheetData, 5, 26, "B", CellValues.String, "2/15/2019 16:48");
                                AddToCell(sheetData, 0, 26, "D", CellValues.String, "Crankcase pressure high");
                                AddToCell(sheetData, 5, 26, "H", CellValues.String, "High");

                                AddToCell(sheetData, 5, 27, "B", CellValues.String, "2/15/2019 10:25");
                                AddToCell(sheetData, 0, 27, "D", CellValues.String, "Generator winding U temperature high");
                                AddToCell(sheetData, 5, 27, "H", CellValues.String, "High");

                            }
                            sheetId++;
                        }


                        excel.Close();
                    }

                    FileStream fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "Report.xlsx", FileMode.Create, FileAccess.Write);
                    memoryStream.WriteTo(fileStream);
                    fileStream.Close();
                    memoryStream.Close();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }


        } // Public Static Void Main Ends

        private static void AddToCell(SheetData sheetData,UInt32Value styleIndex,UInt32 rowIndex, string ColumnName,DocumentFormat.OpenXml.EnumValue<CellValues> CellDataType, string cellValue )
        {
            Row row = new Row() { RowIndex =  rowIndex };
            Cell cell = new Cell();

            cell = new Cell() { StyleIndex = styleIndex};
            cell.CellReference = ColumnName + row.RowIndex.ToString();
            cell.DataType = CellDataType;
            cell.CellValue = new CellValue(cellValue);
            row.AppendChild(cell);

            sheetData.Append(row);
        }

        private static void BuildChart(DrawingsPart dp, string sheetName, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {

            string cnWorksheetName = sheetName;
            // we'll create the chart content first, then create the
            // drawings part of the chart later.
            ChartPart chartp = dp.AddNewPart<ChartPart>();
            chartp.ChartSpace = new C.ChartSpace();

            C.Chart chart = new C.Chart();

            chart.PlotArea = new C.PlotArea();
            chart.PlotArea.Layout = new C.Layout();

            int i, j;
            C.BarChartSeries bcs;

            C.BarChart bc = new C.BarChart();
            bc.BarDirection = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
            bc.BarGrouping = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
           

            C.CategoryAxisData cad = new C.CategoryAxisData();
            cad.StringReference = new C.StringReference();
            //cad.StringReference.Formula = new C.Formula();
            //cad.StringReference.Formula.Text =
            //    string.Format("{0}!${1}${2}:${3}${2}",
            //    sheetName, saColumnName[cnCategoryStartColumnIndex - 1],
            //    cnCategoryRowIndex, saColumnName[cnCategoryStartColumnIndex - 1 + cnDataWidth - 1],
            //    cnCategoryRowIndex);
            cad.StringReference.StringCache = new C.StringCache();
            for (j = 0; j < cnDataWidth; ++j)
            {
                cad.StringReference.StringCache.Append(new C.StringPoint()
                {
                    Index = (uint)j,
                    NumericValue = new C.NumericValue(saCategories[j])
                });
            }
            cad.StringReference.StringCache.PointCount = new C.PointCount() { Val = cnDataWidth };

            

            C.Values vals;
            for (i = 0; i < cnDataHeight; ++i)
            {
                bcs = new C.BarChartSeries();
                bcs.Index = new C.Index() { Val = (uint)i };
                bcs.Order = new C.Order() { Val = (uint)i };
                
                bcs.SeriesText = new C.SeriesText();
                bcs.SeriesText.StringReference = new C.StringReference();
                //bcs.SeriesText.StringReference.Formula = new C.Formula();
                //bcs.SeriesText.StringReference.Formula.Text =
                //    string.Format("{0}!${1}${2}", cnWorksheetName,
                //    saColumnName[(cnCategoryStartColumnIndex - 1) - 1], i + cnDataStartRowIndex);
                bcs.SeriesText.StringReference.StringCache = new C.StringCache();
                bcs.SeriesText.StringReference.StringCache.PointCount = new C.PointCount() { Val = 1 };
                bcs.SeriesText.StringReference.StringCache.Append(new C.StringPoint()
                {
                    Index = 0,
                    NumericValue = new C.NumericValue(saLegend[i])
                });

                

               

                 

                // the contents for the category data is the same for every data series
                // But we can't just append it because the variable is appended by reference
                // and not by value. So we need to clone it.
                bcs.Append((C.CategoryAxisData)cad.CloneNode(true));

                vals = new C.Values();
                vals.NumberReference = new C.NumberReference();
                //vals.NumberReference.Formula = new C.Formula();
                //vals.NumberReference.Formula.Text =
                //    string.Format("{0}!${1}${2}:${3}${2}",
                //    cnWorksheetName, saColumnName[cnCategoryStartColumnIndex - 1],
                //    i + cnDataStartRowIndex, saColumnName[cnCategoryStartColumnIndex - 1 + cnDataWidth - 1],
                //    i + cnDataStartRowIndex);
                vals.NumberReference.NumberingCache = new C.NumberingCache();
                vals.NumberReference.NumberingCache.FormatCode = new C.FormatCode("General");
                for (j = 0; j < cnDataWidth; ++j)
                {
                    vals.NumberReference.NumberingCache.Append(new C.NumericPoint()
                    {
                        Index = (uint)j,
                        NumericValue = new C.NumericValue(faChartData[i, j].ToString())
                    });
                }
                vals.NumberReference.NumberingCache.PointCount = new C.PointCount() { Val = cnDataWidth };
                bcs.Append(vals);

                bc.Append(bcs);
            }

            bc.Append(new C.AxisId() { Val = cnAxisId1 });
            bc.Append(new C.AxisId() { Val = cnAxisId2 });
            bc.Append(new C.Overlap() { Val = -20 });
            bc.Append(new C.GapWidth() { Val = 115 });
            

            chart.PlotArea.Append(bc);

            C.CategoryAxis catax = new C.CategoryAxis();
            catax.AxisId = new C.AxisId() { Val = cnAxisId1 };
            catax.Scaling = new C.Scaling()
            {
                Orientation = new C.Orientation() { Val = C.OrientationValues.MinMax }
            };
            catax.AxisPosition = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
            catax.TickLabelPosition = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
            catax.CrossingAxis = new C.CrossingAxis() { Val = cnAxisId2 };
            catax.Append(new C.Crosses() { Val = C.CrossesValues.AutoZero });
            catax.Append(new C.AutoLabeled() { Val = true });
            catax.Append(new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center });
            catax.Append(new C.LabelOffset() { Val = 100 });
            chart.PlotArea.Append(catax);

            C.ValueAxis valax = new C.ValueAxis();
            valax.AxisId = new C.AxisId() { Val = cnAxisId2 };
            valax.Scaling = new C.Scaling()
            {
                Orientation = new C.Orientation() { Val = C.OrientationValues.MinMax }
            };
            valax.AxisPosition = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
            valax.MajorGridlines = new C.MajorGridlines();
            valax.NumberingFormat = new C.NumberingFormat()
            {
                FormatCode = "General",
                SourceLinked = true
            };
            valax.TickLabelPosition = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
            valax.CrossingAxis = new C.CrossingAxis() { Val = cnAxisId1 };
            valax.Append(new C.Crosses() { Val = C.CrossesValues.AutoZero });
            valax.Append(new C.CrossBetween() { Val = C.CrossBetweenValues.Between });
            chart.PlotArea.Append(valax);

            chart.Legend = new C.Legend();
            chart.Legend.LegendPosition = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
            chart.Legend.Append(new C.Layout());

            chart.PlotVisibleOnly = new C.PlotVisibleOnly() { Val = true };

            chartp.ChartSpace.Append(chart);

            // end of the chart content

            // The drawings part of the chart
            Xdr.GraphicFrame gf = new Xdr.GraphicFrame();
            gf.Macro = string.Empty;
            gf.NonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties();
            // this has to be unique within the WorksheetDrawing class of the DrawingsPart
            // Continue with a different ID for other charts and other images.
            // Yes, normal images too.
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = 2;
            // give a friendly name
            gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = "Chart 1";
            gf.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();

            gf.Transform = new Xdr.Transform();
            gf.Transform.Offset = new A.Offset() { X = 0, Y = 0 };
            gf.Transform.Extents = new A.Extents() { Cx = 0, Cy = 0 };
           

            gf.Graphic = new A.Graphic();
            gf.Graphic.GraphicData = new A.GraphicData();
            gf.Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
            gf.Graphic.GraphicData.Append(new C.ChartReference() { Id = dp.GetIdOfPart(chartp) });

            Xdr.TwoCellAnchor tcanchor = new Xdr.TwoCellAnchor();
            tcanchor.FromMarker = new Xdr.FromMarker();
            tcanchor.FromMarker.RowId = new Xdr.RowId(startRowIndex.ToString());
            // no offset
            tcanchor.FromMarker.RowOffset = new Xdr.RowOffset("0");
            tcanchor.FromMarker.ColumnId = new Xdr.ColumnId(startColumnIndex.ToString());
            // no offset
            tcanchor.FromMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.ToMarker = new Xdr.ToMarker();
            tcanchor.ToMarker.RowId = new Xdr.RowId(endRowIndex.ToString());
            // no offset
            tcanchor.ToMarker.RowOffset = new Xdr.RowOffset("0");
            tcanchor.ToMarker.ColumnId = new Xdr.ColumnId(endColumnIndex.ToString());
            // no offset
            tcanchor.ToMarker.ColumnOffset = new Xdr.ColumnOffset("0");

            tcanchor.Append(gf);
            tcanchor.Append(new Xdr.ClientData());

            dp.WorksheetDrawing.Append(tcanchor);
            dp.WorksheetDrawing.Save();
        }

        private static string[] CalculateExcelColumnName(int NumberOfColumns)
        {
            string[] sa = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            string[] result = new string[NumberOfColumns];
            string s = string.Empty;
            int i, j, k, l;
            i = j = k = -1;
            for (l = 0; l < NumberOfColumns; ++l)
            {
                s = string.Empty;
                ++k;
                if (k == 26)
                {
                    k = 0;
                    ++j;
                    if (j == 26)
                    {
                        j = 0;
                        ++i;
                    }
                }
                if (i >= 0) s += sa[i];
                if (j >= 0) s += sa[j];
                if (k >= 0) s += sa[k];
                result[l] = s;
            }
            return result;
        }
        private static void InsertImage(WorksheetPart sheet1, DrawingsPart drawingsPart2, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, Stream imageStream)
        {
            //Inserting a drawing element in worksheet
            int drawingPartId = GetNextRelationShipID(sheet1);
            GenerateDrawingsPart1Content(drawingsPart2, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
            //GenerateImageDrawingPartContent(drawingsPart2,imageStream);
            //Adding the image
            ImagePart imagePart1 = drawingsPart2.AddNewPart<ImagePart>("image/png", "rId4");
            imagePart1.FeedData(imageStream);
        }


        #region Helper methods
        /// <summary>
        /// Get the index of legacy drawing element in the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Index of legacy drawing</returns>
        private static int GetIndexofLegacyDrawing(WorksheetPart sheet1)
        {
            for (int i = 0; i < sheet1.Worksheet.ChildElements.Count; i++)
            {
                OpenXmlElement element = sheet1.Worksheet.ChildElements[i];
                if (element is LegacyDrawing)
                    return i;
            }
            return -1;
        }
        /// <summary>
        /// Returns the WorksheetPart for the specified sheet name
        /// </summary>
        /// <param name="workbookpart">The WorkbookPart</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>Returns the WorksheetPart for the specified sheet name</returns>
        private static WorksheetPart GetSheetByName(WorkbookPart workbookpart, string sheetName)
        {
            foreach (WorksheetPart sheetPart in workbookpart.WorksheetParts)
            {
                string uri = sheetPart.Uri.ToString();
                if (uri.EndsWith(sheetName + ".xml"))
                    return sheetPart;
            }
            return null;
        }
        /// <summary>
        /// Returns the next relationship id for the specified WorksheetPart
        /// </summary>
        /// <param name="sheet1">The worksheetPart</param>
        /// <returns>Returns the next relationship id </returns>
        private static int GetNextRelationShipID(WorksheetPart sheet1)
        {
            int nextId = 0;
            List<int> ids = new List<int>();
            foreach (IdPartPair part in sheet1.Parts)
            {
                ids.Add(int.Parse(part.RelationshipId.Replace("rId", string.Empty)));
            }
            if (ids.Count > 0)
                nextId = ids.Max() + 1;
            else
                nextId = 1;
            return nextId;
        }

        private static void GenerateImageDrawingPartContent(DrawingsPart drawingsPart, Stream imageStream)
        {
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(imageStream);

            DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
            var extentsCx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
            var extentsCy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
            bm.Dispose();

            var colOffset = 0;
            var rowOffset = 0;
            int colNumber = 4;
            int rowNumber = 4;

            Xdr.WorksheetDrawing worksheetDrawing = new Xdr.WorksheetDrawing();

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0 ?
                (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1 :
                1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = "barcode" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = "rId4", CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }

        // Generates content of drawingsPart1.
        private static void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor() { EditAs = Xdr.EditAsValues.OneCell };

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";// "38100";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = startRowIndex.ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "85725"; //multiply of 9525
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = endRowIndex.ToString();
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "9525";  //multiply of 9525

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 1" };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId4", CompressionState = A.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() {X=0L,Y=0L };// { X = 1257300L, Y = 762000L };
            A.Extents extents1 = new A.Extents() { Cx = 2381250L, Cy = 628650L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            //System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ImageFile);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            //DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();
           /// extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
           // extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
           // bm.Dispose();
          //  transform2D1.Extents = extents;

            //Position pos = new Position();
            //pos.X = 0;
            //pos.Y = 0;
            //Extent ext = new Extent();
            //ext.Cx = extents.Cx;
            //ext.Cy = extents.Cy;
            //AbsoluteAnchor anchor = new AbsoluteAnchor();
            //anchor.Position = pos;
            //anchor.Extent = ext;
            //anchor.Append(picture1);
            //anchor.Append(new ClientData());

            //worksheetDrawing1.Append(anchor);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;


            /*  worksheetDrawing.Append(twoCellAnchor1);*/


            //drawingsPart1.WorksheetDrawing = worksheetDrawing;
        }
        #endregion Helper methods



        private static void InsertPieChartInSpreadSheet(DrawingsPart drawingsPart, string chartTitle, Dictionary<string, int> data, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();


            ChartSpace chartSpace = new ChartSpace();
            chartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());

            ManualLayout manualLayout1 = new ManualLayout();
            LayoutTarget layoutTarget1 = new LayoutTarget() { Val = LayoutTargetValues.Inner };
            LeftMode leftMode1 = new LeftMode() { Val = LayoutModeValues.Edge };
            TopMode topMode1 = new TopMode() { Val = LayoutModeValues.Edge };
            Left left1 = new Left() { Val = 0.5D };
            Top top1 = new Top() { Val = 0.2D };

            Width width1 = new Width() { Val = 0.95622038461448768D };
            Height height1 = new Height() { Val = 0.54928769841269842D };

            manualLayout1.Append(layoutTarget1);
            manualLayout1.Append(leftMode1);
            manualLayout1.Append(topMode1);
            manualLayout1.Append(left1);
            manualLayout1.Append(top1);
            manualLayout1.Append(width1);
            manualLayout1.Append(height1);



            layout.Append(manualLayout1);

            NoFill noFill = new NoFill();
            C.ShapeProperties shapeProperties = new C.ShapeProperties();

            DocumentFormat.OpenXml.Drawing.Outline outline15 = new DocumentFormat.OpenXml.Drawing.Outline();
            DocumentFormat.OpenXml.Drawing.SolidFill noFill17 = new DocumentFormat.OpenXml.Drawing.SolidFill();

            RgbColorModelHex schemeColor29 = new RgbColorModelHex() { Val = "FFFFFF" };

            noFill17.Append(schemeColor29);
            outline15.Append(noFill17);

            shapeProperties.Append(noFill);
            shapeProperties.Append(outline15);
            plotArea.Append(shapeProperties);


            PieChart pieChart = plotArea.AppendChild<PieChart>(new PieChart());

            PieChartSeries pieChartSeries = pieChart.AppendChild<PieChartSeries>(new PieChartSeries(
                new Index() { Val = (UInt32Value)0U },
                new Order() { Val = (UInt32Value)0U },
                new SeriesText(new NumericValue() { Text = "PieChartSeries" })));

            CategoryAxisData catAx = new CategoryAxisData();


            StringReference stringReference = new StringReference();
            StringCache stringCache = new StringCache();

            PointCount pointCount = new PointCount() { Val = (uint)data.Count };

            stringCache.Append(pointCount);

            uint i = 0;
            foreach (var key in data.Keys)
            {
                stringCache.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(key));
                i++;
            }

            stringReference.Append(stringCache);
            catAx.Append(stringReference);
            pieChartSeries.Append(catAx);



            C.Values values = new C.Values();
            NumberReference numberReference = new NumberReference();
            NumberingCache numberingCache = new NumberingCache();

            i = 0;
            foreach (var key in data.Keys)
            {
                numberingCache.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(data[key].ToString()));
                i++;
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);
            pieChartSeries.Append(values);

            AddChartTitle(chart, chartTitle);
            pieChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            pieChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });


            CategoryAxis catAx1 = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
                 new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
                 new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
                 new CrossingAxis() { Val = new UInt32Value(48672768U) },
                 new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                 new AutoLabeled() { Val = new BooleanValue(true) },
                 new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
                 new LabelOffset() { Val = new UInt16Value((ushort)100) }));




            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                    DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                }),
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                }, new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>
            (TickLabelPositionValues.NextTo)
                }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));

            // Add the chart Legend.
            Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Bottom) },
                new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            chartPart.ChartSpace = chartSpace;

            PositionChart(chartPart, drawingsPart, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
        }

        private static void PositionChart(ChartPart chartPart, DrawingsPart drawingsPart, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(startColumnIndex.ToString()),
                                            new ColumnOffset("581025"),
                                            new RowId(startRowIndex.ToString()),
                                            new RowOffset("114300")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId(endColumnIndex.ToString()),
                new ColumnOffset("276225"),
                new RowId(endRowIndex.ToString()),
                new RowOffset("0")));

            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                    new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());
        }

        private static void AddChartTitle(DocumentFormat.OpenXml.Drawing.Charts.Chart chart, string title)
        {
            var ctitle = chart.AppendChild(new Title());
            var chartText = ctitle.AppendChild(new ChartText());
            var richText = chartText.AppendChild(new RichText());

            var bodyPr = richText.AppendChild(new BodyProperties());
            var lstStyle = richText.AppendChild(new ListStyle());
            var paragraph = richText.AppendChild(new Paragraph());

            var apPr = paragraph.AppendChild(new ParagraphProperties());
            apPr.AppendChild(new DefaultRunProperties());

            var run = paragraph.AppendChild(new DocumentFormat.OpenXml.Drawing.Run());
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.RunProperties() { Language = "en-CA" });
            run.AppendChild(new DocumentFormat.OpenXml.Drawing.Text() { Text = title });
            //ctitle.AppendChild(new Overlay() { Val = new BooleanValue(false) });

        }


        // Given a document name, a worksheet name, a chart title, and a Dictionary collection of text keys
        // and corresponding integer data, creates a column chart with the text as the series and the integers as the values.
        private static void InsertBarChartInSpreadsheet(Worksheet ws, DrawingsPart drawingsPart, string title, Dictionary<string, int> data, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
        {
            WorksheetPart wsp = ws.WorksheetPart;
            // Open the document for editing.
            //    using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            //     {
            //        IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().
            //Where(s => s.Name == worksheetName);
            //        if (sheets.Count() == 0)
            //        {
            //            // The specified worksheet does not exist.
            //            return;
            //        }
            //        WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);

            // Add a new drawing to the worksheet.
            //DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            //worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
            //{ Id = worksheetPart.GetIdOfPart(drawingsPart) });
            //worksheetPart.Worksheet.Save();

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart = plotArea.AppendChild<BarChart>(new BarChart(new BarDirection()
            { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) }));

            uint i = 0;

            barChart.Append(new C.Overlap() { Val = -100 });
            barChart.Append(new C.GapWidth() { Val = 219 });
            barChart.Append(new C.VaryColors() { Val = false });

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.

            BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
            {
                Val = (UInt32Value)0U
            },
               new Order() { Val = (UInt32Value) 0U},
               new SeriesText(new NumericValue() { Text = "Test" })));


            //foreach (DateTime key in data.Keys)
            //{


            //    StringLiteral strLit = barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral());
            //    //3strLit.Append(new FormatCode("dd/mm/yyyy;@"));
            //    strLit.Append(new PointCount() { Val = new UInt32Value(i) });
            //    strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(key.Date.Ticks.ToString()));

            //    NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
            //        new DocumentFormat.OpenXml.Drawing.Charts.Values()).AppendChild<NumberLiteral>(new NumberLiteral());
            //    numLit.Append(new FormatCode("General"));
            //    numLit.Append(new PointCount() { Val = new UInt32Value(1U) });
            //    numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(0u) }).Append(new NumericValue(data[key].ToString()));

            //    i++;
            //}

            CategoryAxisData catAxData = new CategoryAxisData();


            StringReference stringReference = new StringReference();
            StringCache stringCache = new StringCache();

            PointCount pointCount = new PointCount() { Val = (uint)data.Count };

            stringCache.Append(pointCount);

            foreach (var key in data.Keys)
            {
                stringCache.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(key));
                i++;
            }

            stringReference.Append(stringCache);
            catAxData.Append(stringReference);
            barChartSeries.Append(catAxData);



            C.Values values = new C.Values();
            NumberReference numberReference = new NumberReference();
            NumberingCache numberingCache = new NumberingCache();

            i = 0;
            foreach (var key in data.Keys)
            {
                numberingCache.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) }).Append(new NumericValue(data[key].ToString()));
                i++;
            }

            numberReference.Append(numberingCache);
            values.Append(numberReference);
            barChartSeries.Append(values);

            AddChartTitle(chart, title);

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            // Add the Category Axis.
            //CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            //{ Val = new UInt32Value(48650112u) }, 
            //new Scaling(new Orientation()
            //{
            //    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            //}),
            //    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
            //    new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
            //    new CrossingAxis() { Val = new UInt32Value(48672768U) },
            //    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //    new AutoLabeled() { Val = new BooleanValue(true) },
            //    new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
            //    new LabelOffset() { Val = new UInt16Value((ushort)100) }));


            CategoryAxis catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis(new AxisId()
            { Val = new UInt32Value(48650112u) }, new Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation()
            {
                Val = new EnumValue<DocumentFormat.
OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }),
           //START NEW CODE
           new Delete() { Val = false },
           //END NEW CODE
           new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
           //START NEW CODE
           new NumberingFormat() { FormatCode = "dd/mm/yyyy;@", SourceLinked = true },
           new MajorTickMark() { Val = TickMarkValues.Outside },
           new MinorTickMark() { Val = TickMarkValues.Cross },
           //END NEW CODE
           new TickLabelPosition() { Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo) },
           new CrossingAxis() { Val = new UInt32Value(48650112u) },
           new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
           //new AutoLabeled() { Val = new BooleanValue(true) },
           new LabelAlignment() { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) },
           new LabelOffset() { Val = new UInt16Value((ushort)100) },
           //START NEW CODE
           new NoMultiLevelLabels() { Val = true }
           //new C.TextProperties(new BodyProperties { Rotation = -5400000, Vertical = TextVerticalValues.Horizontal },
           //         new ListStyle(),
           //         new Paragraph(new ParagraphProperties(new DefaultRunProperties(), new EndParagraphRunProperties { Language = "en-US" })))
           //END NEW CODE
           ));
            // Add the Date Axis.
            //DateAxis dateAx = plotArea.AppendChild<DateAxis>(new DateAxis(new AxisId() { Val = new UInt32Value(48650112u) },
            //    new Scaling(new Orientation()
            //    {
            //        Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax),

            //    }, new MaxAxisValue() { Val = 1020D }, new MinAxisValue() { Val = 840D }),
            //    new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) },
            //    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
            //    {
            //        FormatCode = "dd/mm/yyyy;@",
            //        SourceLinked = new BooleanValue(true)
            //    },
            //    //new MajorTickMark() { Val = TickMarkValues.None },
            //    //new MinorTickMark() { Val = TickMarkValues.None },
            //    new TickLabelPosition()
            //    {
            //        Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
            //    },
            //    //new MajorGridlines(),

            //    new CrossingAxis() { Val = new UInt32Value(48650112U) },
            //    new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
            //    new AutoLabeled() { Val=true},
            //    new LabelOffset() { Val = (UInt16Value)100U },
            //    //new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) },
            //    new BaseTimeUnit() { Val= TimeUnitValues.Days }));

            //C.DateAxis dateAxis3 = new C.DateAxis();
            //C.AxisId axisId15 = new C.AxisId() { Val = (UInt32Value)48650112u };

            //C.Scaling scaling7 = new C.Scaling();
            //C.Orientation orientation7 = new C.Orientation() { Val = C.OrientationValues.MinMax };

            //scaling7.Append(orientation7);
            //C.Delete delete7 = new C.Delete() { Val = false };
            //C.AxisPosition axisPosition7 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
            //C.NumberingFormat numberingFormat8 = new C.NumberingFormat() { FormatCode = "dd/mm/yyyy;@", SourceLinked = false };
            //C.MajorTickMark majorTickMark7 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
            //C.MinorTickMark minorTickMark7 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
            //C.TickLabelPosition tickLabelPosition7 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

            //C.ChartShapeProperties chartShapeProperties37 = new C.ChartShapeProperties();
            //A.NoFill noFill78 = new A.NoFill();

            //A.Outline outline145 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            //A.SolidFill solidFill177 = new A.SolidFill();

            //A.SchemeColor schemeColor361 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            //A.LuminanceModulation luminanceModulation203 = new A.LuminanceModulation() { Val = 15000 };
            //A.LuminanceOffset luminanceOffset178 = new A.LuminanceOffset() { Val = 85000 };

            //schemeColor361.Append(luminanceModulation203);
            //schemeColor361.Append(luminanceOffset178);

            //solidFill177.Append(schemeColor361);
            //A.Round round84 = new A.Round();

            //outline145.Append(solidFill177);
            //outline145.Append(round84);
            //A.EffectList effectList61 = new A.EffectList();

            //chartShapeProperties37.Append(noFill78);
            //chartShapeProperties37.Append(outline145);
            //chartShapeProperties37.Append(effectList61);

            //C.TextProperties textProperties23 = new C.TextProperties();
            //A.BodyProperties bodyProperties36 = new A.BodyProperties() { Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true };
            //A.ListStyle listStyle36 = new A.ListStyle();

            //A.Paragraph paragraph36 = new A.Paragraph();

            //A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties();

            //A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties() { FontSize = 900, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0 };

            //A.SolidFill solidFill178 = new A.SolidFill();

            //A.SchemeColor schemeColor362 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            //A.LuminanceModulation luminanceModulation204 = new A.LuminanceModulation() { Val = 65000 };
            //A.LuminanceOffset luminanceOffset179 = new A.LuminanceOffset() { Val = 35000 };

            //schemeColor362.Append(luminanceModulation204);
            //schemeColor362.Append(luminanceOffset179);

            //solidFill178.Append(schemeColor362);
            //A.LatinFont latinFont28 = new A.LatinFont() { Typeface = "+mn-lt" };
            //A.EastAsianFont eastAsianFont28 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            //A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            //defaultRunProperties29.Append(solidFill178);
            //defaultRunProperties29.Append(latinFont28);
            //defaultRunProperties29.Append(eastAsianFont28);
            //defaultRunProperties29.Append(complexScriptFont28);

            //paragraphProperties32.Append(defaultRunProperties29);
            //A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties() { Language = "en-US" };

            //paragraph36.Append(paragraphProperties32);
            //paragraph36.Append(endParagraphRunProperties29);

            //textProperties23.Append(bodyProperties36);
            //textProperties23.Append(listStyle36);
            //textProperties23.Append(paragraph36);
            //C.CrossingAxis crossingAxis7 = new C.CrossingAxis() { Val = (UInt32Value)48650112U };
            //C.Crosses crosses7 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
            //C.AutoLabeled autoLabeled4 = new C.AutoLabeled() { Val = true };
            //C.LabelOffset labelOffset4 = new C.LabelOffset() { Val = (UInt16Value)100U };
            //C.BaseTimeUnit baseTimeUnit3 = new C.BaseTimeUnit() { Val = C.TimeUnitValues.Days };

            //dateAxis3.Append(axisId15);
            //dateAxis3.Append(scaling7);
            //dateAxis3.Append(delete7);
            //dateAxis3.Append(axisPosition7);
            //dateAxis3.Append(numberingFormat8);
            //dateAxis3.Append(majorTickMark7);
            //dateAxis3.Append(minorTickMark7);
            //dateAxis3.Append(tickLabelPosition7);
            //dateAxis3.Append(chartShapeProperties37);
            //dateAxis3.Append(textProperties23);
            //dateAxis3.Append(crossingAxis7);
            //dateAxis3.Append(crosses7);
            //dateAxis3.Append(autoLabeled4);
            //dateAxis3.Append(labelOffset4);
            //dateAxis3.Append(baseTimeUnit3);

            //plotArea.AppendChild(dateAxis3);


            // Add the Value Axis.
            ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() { Val = new UInt32Value(48672768u) },
                new Scaling(new Orientation()
                {
                    Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax),

                }, new MaxAxisValue() { Val = 1020D }, new MinAxisValue() { Val = 840D }),
                 new Delete() { Val = false },
                new AxisPosition() { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left) },
                new MajorGridlines(),
                new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                {
                    FormatCode = new StringValue("General"),
                    SourceLinked = new BooleanValue(true)
                }, new TickLabelPosition()
                {
                    Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
                }, new CrossingAxis() { Val = new UInt32Value(48650112U) },
                new Crosses() { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) },
                new CrossBetween() { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) }));


            // Add the chart Legend.
            //  Legend legend = chart.AppendChild<Legend>(new Legend(new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Bottom) },new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // Position the chart on the worksheet using a TwoCellAnchor object.
            //drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(startColumnIndex.ToString()),
                new ColumnOffset("0"),
                new RowId(startRowIndex.ToString()),
                new RowOffset("0")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId(endColumnIndex.ToString()),
                new ColumnOffset("0"),
                new RowId(endRowIndex.ToString()),
                new RowOffset("0")));

            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
    Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
    Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties() { Id = new UInt32Value(2u), Name = "Chart 1" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                                                                    new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            twoCellAnchor.Append(new ClientData());

            WorksheetDrawing wsd;

            if (wsp.DrawingsPart != null)
            {
                wsd = drawingsPart.WorksheetDrawing;
                //wsd.Append(twoCellAnchor);
            }


            //    // Save the WorksheetDrawing object.
            //    drawingsPart.WorksheetDrawing.Save();
            //}

        }

        private static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
            new DocumentFormat.OpenXml.Spreadsheet.Fonts(
            new DocumentFormat.OpenXml.Spreadsheet.Font(new FontSize() { Val = 11 },new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },new FontName() { Val = "Calibri" }),// Index 0 - The default font.
            new Font(new Bold(),new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 1 - The bold font.
            new Font(new Italic(),new FontSize() { Val = 11 },new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 2 - The Italic font.
            new Font(new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },new FontName() { Val = "Calibri" }),  // Index 3 - The Times Roman font. with 16 size
            new Font(new Bold(),new FontSize() { Val = 18 }, new Color() { Rgb = new HexBinaryValue() { Value = "000000" } }, new FontName() { Val = "Calibri" }),  // Index 4 - The Times Roman font. with 16 size
            new Font(new Bold(), new FontSize() { Val = 11 }, new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } }, new FontName() { Val = "Calibri" })  // Index 5 - The bold font.

            ),
            new Fills(
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 0 - The default fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.None }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 1 - The default fill of gray 125 (required)
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill() { PatternType = PatternValues.Gray125 }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 2 - The yellow fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
            )
            { PatternType = PatternValues.Solid }),
            new DocumentFormat.OpenXml.Spreadsheet.Fill( // Index 3 - The Blue fill.
            new DocumentFormat.OpenXml.Spreadsheet.PatternFill(
            new DocumentFormat.OpenXml.Spreadsheet.ForegroundColor() { Rgb = new HexBinaryValue() { Value = "8EA9DB" } }
            )
            { PatternType = PatternValues.Solid })
            ),
            new Borders(
            new Border( // Index 0 - The default border.
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(),
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(),
            new DiagonalBorder()),
            new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder()),
                   new Border( // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            new DocumentFormat.OpenXml.Spreadsheet.LeftBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.RightBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.TopBorder(
            new Color() { Auto = true }
            )
            { Style = BorderStyleValues.None },
            new DocumentFormat.OpenXml.Spreadsheet.BottomBorder(
            new Color() { Rgb = new HexBinaryValue() { Value = "70AD47" } }
            )
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder())
            ),
            new CellFormats(
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }, // Index 0 - The default cell style. If a cell does not have a style index applied it will use this style combination instead
            new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 1 - Bold
            new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 2 - Italic
            new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 3 - Times Roman
            new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true }, // Index 4 - Yellow Fill
            new CellFormat( // Index 5 - Alignment
            new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            )
            { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },
            new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }, // Index 6 - Border
             new CellFormat( new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 7 - Alignment
             { FontId = 1, FillId = 0, BorderId = 0, ApplyAlignment = true },

             new CellFormat() { FontId = 4, FillId = 0, BorderId = 0, ApplyFont = true }, // Index 8 - Times Roman
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 2, ApplyFont = true }, // Index 9 - Bottom Border with Color 70AD47
             new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) // Index 10 - Alignment
             { FontId =5, FillId = 3, BorderId = 0, ApplyAlignment = true }


             )
            ); // return
        }
    }
}
