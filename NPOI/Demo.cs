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
              .SetFontColor(IndexedColors.White)
              .SetBold(true)
              .SetBackground(FillPattern.SolidForeground, Color.Purple)
              .Style;

            sheet.ColumnWidth(0, 30);
            sheet.ColumnWidth(1, 50);
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
                    sheet.AddCell(statHeader, rowStart, 1, $"{stat.JobName} {stat.Started}");
                    sheet.AddCell(statHeader, rowStart, 2, String.Empty);
                }
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
