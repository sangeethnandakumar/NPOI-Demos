using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Drawing;

namespace NPOI.Demo
{
    public class Statitics
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public Guid CatageoryId { get; set; }
        public string JobName { get; set; }
        public string CatageoryName { get; set; }
        public DateTime Date { get; set; }
        public DateTime Started { get; set; }
        public DateTime Ended { get; set; }
        public TimeSpan Duration { get { return (Ended - Started).Duration(); } }
        public int AmazonStartPage { get; set; }
        public int AmazonEndPage { get; set; }
        public int ScrapHeroQuotasUsed { get; set; }
        public int TotalScrapped { get; set; }
        public int Duplicates { get; set; }
        public int Processed { get { return TotalScrapped - Duplicates; } }
        public int Success { get; set; }
        public int NonUsefull { get { return Blacklisted + Above20MBSize + Above150KRank + ScrapHeroError + BookNotInZLib; } }
        public int Blacklisted { get; set; }
        public int Above20MBSize { get; set; }
        public int Above150KRank { get; set; }
        public int ScrapHeroError { get; set; }
        public int BookNotInZLib { get; set; }
    }

    public static class Demo
    {
        public static void MakeExcelFile(List<Statitics> statitics)
        {
            //WorkBook
            var workbook = new XSSFWorkbook();
            //Sheet
            var sheet = workbook.CreateSheet("Statitics");
            //Style
            var catageoryHeader = new ExcelStyle(workbook)
                .SetFont("Calibri")
                .SetFontSize(12)
                .SetAllignment(HorizontalAlignment.Center, VerticalAlignment.Center)
                .SetFontColor(IndexedColors.Black)
                .SetBold(true)
                .SetBackground(FillPattern.SolidForeground, Color.Yellow)
                .Style;

            var statHeader = new ExcelStyle(workbook)
              .SetFont("Calibri")
              .SetFontSize(12)
              .SetAllignment(HorizontalAlignment.Left, VerticalAlignment.Center)
              .SetFontColor(IndexedColors.Black)
              .SetBold(true)
              .SetBackground(FillPattern.SolidForeground, Color.FromArgb(217, 225, 242))
              .Style;

            var defaultStyle = new ExcelStyle(workbook)
              .SetFont("Calibri")
              .SetFontSize(12)
              .SetAllignment(HorizontalAlignment.Left, VerticalAlignment.Center)
              .SetFontColor(IndexedColors.Black)
              .SetBold(true)
              .SetBackground(FillPattern.NoFill, Color.White)
              .Style;

            var primaryStyle = new ExcelStyle(workbook)
             .SetFont("Calibri")
             .SetFontSize(12)
             .SetAllignment(HorizontalAlignment.Left, VerticalAlignment.Center)
             .SetFontColor(IndexedColors.Black)
             .SetBold(true)
             .SetBackground(FillPattern.SolidForeground, Color.FromArgb(255, 230, 153))
             .Style;

            var secondaryStyle = new ExcelStyle(workbook)
                .SetFont("Calibri")
                .SetFontSize(12)
                .SetAllignment(HorizontalAlignment.Left, VerticalAlignment.Center)
                .SetFontColor(IndexedColors.Black)
                .SetBold(true)
                .SetBackground(FillPattern.Squares, Color.FromArgb(255, 230, 153))
                .Style;

            var successStyle = new ExcelStyle(workbook).DefaultSuccess().Style;
            var warningStyle = new ExcelStyle(workbook).DefaultWarning().Style;
            var errorsStyle = new ExcelStyle(workbook).DefaultError().Style;

            sheet.ColumnWidth(0, 30);
            sheet.ColumnWidth(1, 40);
            sheet.ColumnWidth(2, 50);

            //Data
            int rowStart = 1;
            //Group by catageory
            var statiticsGrouped = statitics.GroupBy(x => x.CatageoryName).ToList();
            foreach (var group in statiticsGrouped)
            {
                sheet.AddRow(catageoryHeader, rowStart, group.Key);
                foreach (var stat in group)
                {
                    sheet.AddCell(statHeader, rowStart, 1, $"{stat.JobName} - [{stat.Started}]");
                    sheet.AddCell(statHeader, rowStart, 2, String.Empty);
                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Started", stat.Started.ToString());
                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Ended", stat.Ended.ToString());
                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Duration", stat.Duration.ToString());
                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Amazon Pages reached", $"{stat.AmazonStartPage} ⟶ {stat.AmazonEndPage}");
                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "ScrapHero Quotas Used", stat.ScrapHeroQuotasUsed.ToString());

                    sheet.AddRow(secondaryStyle, ++rowStart, String.Empty, "Total Scrapped", stat.TotalScrapped.ToString());
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;

                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Duplicates", stat.Duplicates.ToString());
                    sheet.AddRow(primaryStyle, ++rowStart, String.Empty, "Processed", $"{stat.Processed} ({stat.TotalScrapped} - {stat.Duplicates})");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;

                    sheet.AddRow(successStyle, ++rowStart, String.Empty, "         ⟶ [SUCCESS]", $"{stat.Success} (out of {stat.Processed})");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;
                    sheet.GetRow(rowStart).GetCell(2).CellStyle = defaultStyle;

                    sheet.AddRow(defaultStyle, ++rowStart, String.Empty, "Total Non-Usefull books (Breakdown)", $"{stat.NonUsefull} books");

                    sheet.AddRow(warningStyle, ++rowStart, String.Empty, "         ⟶ [GENERE BLACKLISTED]", $"{stat.Blacklisted} books");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;
                    sheet.AddRow(warningStyle, ++rowStart, String.Empty, "         ⟶ [EPUB ABOVE 20MB]", $"{stat.Above20MBSize} books");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;
                    sheet.AddRow(warningStyle, ++rowStart, String.Empty, "         ⟶ [EPUB RANK ABOVE 150K]", $"{stat.Above150KRank} books");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;

                    sheet.AddRow(errorsStyle, ++rowStart, String.Empty, "         ⟶ [SCRAPHERO ERROR]", $"{stat.ScrapHeroError} books");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;
                    sheet.AddRow(errorsStyle, ++rowStart, String.Empty, "         ⟶ [BOOK NOT IN ZLIB]", $"{stat.BookNotInZLib} books");
                    sheet.GetRow(rowStart).GetCell(0).CellStyle = defaultStyle;


                    var selection2 = sheet.BoxSelection(rowStart - 5, 1, rowStart-1, 1);
                    sheet.DrawBorderOnSelection(selection2, BorderCover.AroundSelectionBox, BorderStyle.Double);

                    var selection = sheet.BoxSelection(rowStart - 15, 1, rowStart-1, 1);
                    sheet.DrawBorderOnSelection(selection, BorderCover.AroundSelectionBox, BorderStyle.DashDot);

                    rowStart += 2;

                    //Set border to selection
                    //var selecction2 = sheet.BoxSelection("B11", "c17");


                }

                //Writing
                var xfile = new FileStream("demo1.xlsx", FileMode.Create, FileAccess.Write);
                workbook.Write(xfile);
                workbook.Close();
                xfile.Close();
                //Opening
                new Process
                {
                    StartInfo = new ProcessStartInfo(@"demo1.xlsx")
                    {
                        UseShellExecute = true
                    }
                }.Start();
            }
        }
    }
}