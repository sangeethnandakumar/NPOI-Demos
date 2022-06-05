using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Drawing;

namespace NPOI.Demo
{

    public class ExcelStyle
    {
        private readonly XSSFWorkbook workbook;
        private IFont font;
        public XSSFCellStyle Style { get; set; }

        const string DEFAULT_FONT = "Tahoma";

        public ExcelStyle(XSSFWorkbook workbook)
        {
            this.workbook = workbook;
            Style = (XSSFCellStyle)workbook.CreateCellStyle();
            Style.IsLocked = false;
            font = workbook.CreateFont();
            SetFont(DEFAULT_FONT);
            SetFontSize(10);
        }

        public ExcelStyle DefaultSuccess()
        {
            return new ExcelStyle(workbook)
                        .SetFontColor(IndexedColors.Green)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.LightGreen);
        }

        public ExcelStyle DefaultWarning()
        {
            return new ExcelStyle(workbook)
                        .SetFontColor(IndexedColors.Black)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.LightYellow);
        }

        public ExcelStyle DefaultError()
        {
            return new ExcelStyle(workbook)
                        .SetFontColor(IndexedColors.Red)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.DarkRed);
        }

        public ExcelStyle SetFont(string fontName)
        {
            font.FontName = fontName;
            Style.SetFont(font);
            return this;
        }

        public ExcelStyle SetFontSize(short fontSize)
        {
            font.FontHeightInPoints = fontSize;
            Style.SetFont(font);
            return this;
        }

        public ExcelStyle SetFontColor(IndexedColors color)
        {
            font.Color = color.Index;
            Style.SetFont(font);
            return this;
        }        

        public ExcelStyle SetBold(bool isBold)
        {
            font.IsBold = isBold;
            Style.SetFont(font);
            return this;
        }

        public ExcelStyle WrapText()
        {
            Style.WrapText = true;
            return this;
        }

        public ExcelStyle SetAllignment(HorizontalAlignment horizontal, VerticalAlignment vertical)
        {
            Style.Alignment = horizontal;
            Style.VerticalAlignment = vertical;
            return this;
        }

        public ExcelStyle SetBackground(FillPattern pattern, IndexedColors color)
        {
            Style.FillPattern = pattern;
            Style.FillForegroundColor = color.Index;
            return this;
        }

        public ExcelStyle SetBackground(FillPattern pattern, Color color)
        {
            Style.FillPattern = pattern;
            Style.FillForegroundXSSFColor = new XSSFColor(color);
            return this;
        }


        public ExcelStyle SetBorder(BorderStyle left, BorderStyle top, BorderStyle right, BorderStyle bottom)
        {
            Style.BorderLeft = left;
            Style.BorderTop = top;
            Style.BorderRight = right;
            Style.BorderBottom = bottom;
            return this;
        }

        public ExcelStyle SetBorderSize(BorderStyle left, BorderStyle top, BorderStyle right, BorderStyle bottom)
        {
            Style.BorderLeft = left;
            Style.BorderTop = top;
            Style.BorderRight = right;
            Style.BorderBottom = bottom;
            return this;
        }

        public static XSSFColor FromColor(Color color)
        {
            return new XSSFColor(color);
        }
    }
}
