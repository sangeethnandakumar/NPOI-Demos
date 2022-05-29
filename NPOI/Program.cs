
using NPOI.Demo;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.CreateSheet("Sheet A");


//Build global styles
var style = new ExcelStyle(workbook)
                        .SetFontColor(IndexedColors.Black)
                        .SetBold(true)
                        .SetBackground(FillPattern.Diamonds, ExcelStyle.IndexedColor(IndexedColors.Green))
                        .Style;

var hyperlink = new ExcelStyle(workbook)
                        .SetFontColor(IndexedColors.Blue)
                        .SetBold(true)
                        .SetBackground(FillPattern.SolidForeground, ExcelStyle.IndexedColor(IndexedColors.White))
                        .Style;

//Defaul default styled and custom styled row values
sheet.AddRow(0, "Sangeeth", "Nandakumar", "RREXS");
sheet.AddRow(style, 1, "Navaneeth", "Nandakumar", "PPXE");
sheet.AddRow(2, "Surya", "Nandakumar");
sheet.AddRow(style, 3, "K", "Nandakumar");

//Set col size
sheet.ColumnWidth(0, 10);
sheet.ColumnWidth(1, 20);
sheet.ColumnWidth(2, 30);

//Specific style for cell
var cell = sheet.GetRow(0).GetCell(1);
cell.CellStyle = hyperlink;

//Making a cell hyperlink
cell.Hyperlink = new XSSFHyperlink(HyperlinkType.Url)
{
    Address = "https://www.google.com"
};

//Load image to workbook
byte[] data = File.ReadAllBytes(@"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\copilot.png");
var imageIndex = workbook.LoadImage(data, PictureType.PNG);

//Insert image from workbook by image index
sheet.InsertImage(workbook, imageIndex , 0, 1);


//Preset usefull styles
//SUCCESS
sheet.GetRow(0).GetCell(0).CellStyle = new ExcelStyle(workbook).DefaultSuccess().Style;
sheet.GetRow(0).GetCell(0).SetCellValue("Success");
//SUCCESS
sheet.GetRow(0).GetCell(1).CellStyle = new ExcelStyle(workbook).DefaultWarning().Style;
sheet.GetRow(0).GetCell(1).SetCellValue("Warning");
//SUCCESS
sheet.GetRow(0).GetCell(2).CellStyle = new ExcelStyle(workbook).DefaultError().Style;
sheet.GetRow(0).GetCell(2).SetCellValue("Error");


//Increase first row height
var firstRow = sheet.GetRow(0);
firstRow.Height = 30 * 50;

//Center text & set border
var centerStyle = new ExcelStyle(workbook)
                        .SetAllignment(HorizontalAlignment.Center, VerticalAlignment.Center) 
                        .WrapText()
                        .SetBold(true)
                        .SetFontSize(20)
                        .SetBorder(BorderStyle.Thin, BorderStyle.Thick, BorderStyle.SlantedDashDot, BorderStyle.Dotted)
                        .Style;
firstRow.GetCell(0).CellStyle = centerStyle;



var xfile = new FileStream("test.xlsx", FileMode.Create, FileAccess.Write);
workbook.Write(xfile);
workbook.Close();
xfile.Close();