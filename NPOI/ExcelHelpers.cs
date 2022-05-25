using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Demo
{
    public static class ExcelHelpers
    {
        public static void AddRow(this ISheet sheet, XSSFCellStyle style, int rowNum,  params string[] values)
        {
            var row = sheet.CreateRow(rowNum);
            for (var i=0; i<values.Length; i++)
            {
                var cell = row.CreateCell(i);
                cell.CellStyle = style;                
                cell.SetCellValue(values[i]);                   
            }
        }
        public static void AddRow(this ISheet sheet, int rowNum, params string[] values)
        {
            var row = sheet.CreateRow(rowNum);
            for (var i = 0; i < values.Length; i++)
            {
                row.CreateCell(i).SetCellValue(values[i]);
            }
        }

        public static int LoadImage(this XSSFWorkbook workbook, byte[] image, PictureType format)
        {
            return workbook.AddPicture(image, format);
        }            
            public static void InsertImage(this ISheet sheet, XSSFWorkbook workbook, int imageIndex, int row, int column)
        {
           
            ICreationHelper helper = workbook.GetCreationHelper();
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = helper.CreateClientAnchor();
            anchor.Col1 = column;
            anchor.Row1 = row;
            IPicture picture = drawing.CreatePicture(anchor, imageIndex);
            picture.Resize();
        }        


        public static void ColumnWidth(this ISheet sheet, int columnNum, int size)
        {
            sheet.SetColumnWidth(columnNum, 256*size);
        }
    }
}
