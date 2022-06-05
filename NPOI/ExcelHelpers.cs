using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace NPOI.Demo
{
    public enum BorderCover
    {
        AroundSelectionBox
    }

    public static class ExcelHelpers
    {

        public static T Clone<T>(this T source)
        {
            var serialized = JsonConvert.SerializeObject(source);
            return JsonConvert.DeserializeObject<T>(serialized);
        }

        public static void DrawBorderOnSelection(this ISheet sheet, Tuple<CellReference, CellReference> range, BorderCover borderCover, BorderStyle borderStyle)
        {
            for (var i = range.Item1.Row; i < range.Item2.Row+1; i++)
            {
                var row = sheet.GetRow(i);
                if(row is not null)
                {

                    for (var j = range.Item1.Col; j < range.Item2.Col + 1; j++)
                    {
                        XSSFCellStyle clonedStyle;
                        var cell = row.GetCell(j);

                        var currentCellStyle = cell?.CellStyle as XSSFCellStyle;
                        if (currentCellStyle is null)
                        {
                            clonedStyle = new ExcelStyle((XSSFWorkbook)sheet.Workbook).DefaultSuccess().Style;
                        }
                        else
                        {
                            clonedStyle = currentCellStyle.Clone() as XSSFCellStyle;
                        }


                        //Outerbox
                        if (borderCover == BorderCover.AroundSelectionBox)
                        {
                            if (i == range.Item1.Row && j == range.Item1.Col)
                            {
                                clonedStyle.BorderLeft = borderStyle;
                                clonedStyle.BorderTop = borderStyle;
                            }
                            else if (j == range.Item2.Col && i == range.Item1.Row)
                            {
                                clonedStyle.BorderTop = borderStyle;
                                clonedStyle.BorderRight = borderStyle;
                            }
                            else if (i == range.Item1.Row)
                            {
                                clonedStyle.BorderTop = borderStyle;
                            }


                            if (i == range.Item2.Row && j == range.Item2.Col)
                            {
                                clonedStyle.BorderBottom = borderStyle;
                            }
                            else if (j == range.Item1.Col && i == range.Item2.Row)
                            {
                                clonedStyle.BorderBottom = borderStyle;
                            }
                            else if (i == range.Item2.Row)
                            {
                                clonedStyle.BorderBottom = borderStyle;
                            }


                            if (j == range.Item1.Col)
                            {
                                clonedStyle.BorderLeft = borderStyle;
                            }

                            if (j == range.Item2.Col)
                            {
                                clonedStyle.BorderRight = borderStyle;
                            }
                        }



                        if (cell is null)
                        {
                            row.CreateCell(j);
                            cell = row.GetCell(j);
                        }
                        cell.CellStyle = clonedStyle;
                    }
                }

            }
        }

        public static Tuple<CellReference, CellReference> BoxSelection(this ISheet sheet, int startRow, int startColumn, int rowOffset, int columnOffset)
        {
            var startReference = new CellReference(startRow, startColumn);
            var endReference = new CellReference(startRow + rowOffset, startColumn + columnOffset);
            return new Tuple<CellReference, CellReference>(startReference, endReference);
        }

        public static Tuple<CellReference, CellReference> BoxSelection(this ISheet sheet, string startAddress, string endAddress)
        {
            var startReference = new CellReference(startAddress);
            var endReference = new CellReference(endAddress);
            return new Tuple<CellReference, CellReference>(startReference, endReference);
        }        


        public static void AddRow(this ISheet sheet, XSSFCellStyle style, int rowNum, params string[] values)
        {
            var row = sheet.CreateRow(rowNum);
            for (var i = 0; i < values.Length; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(values[i]);
                cell.CellStyle = style;
            }
        }

        public static void AddCell(this ISheet sheet, XSSFCellStyle style, int rowNum, int colNum, string value)
        {
            AddCell(sheet, rowNum, colNum, value);
            sheet.GetRow(rowNum).GetCell(colNum).CellStyle = style;
        }

        public static void AddCell(this ISheet sheet, int rowNum, int colNum, string value)
        {
            var row = sheet.GetRow(rowNum);
            if(row == null)
            {
                row = sheet.CreateRow(rowNum);
            }       
            row.CreateCell(colNum).SetCellValue(value);
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
            sheet.SetColumnWidth(columnNum, 256 * size);
        }

    }
}
