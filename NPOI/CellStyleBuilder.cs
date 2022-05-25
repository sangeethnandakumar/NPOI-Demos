using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Demo
{

    public class CellStyleBuilder
    {
        private readonly XSSFWorkbook workbook;
        private IFont font;
        public XSSFCellStyle Style { get; set; }

        public CellStyleBuilder(XSSFWorkbook workbook)
        {
            this.workbook = workbook;
            Style = (XSSFCellStyle)workbook.CreateCellStyle();
            Style.IsLocked = false;
            Style.WrapText = true;
        }

        public CellStyleBuilder DefaultSuccess()
        {
            return new CellStyleBuilder(workbook)
                        .SetFont("Areal")
                        .SetFontColor(IndexedColors.Green)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.LightGreen);
        }

        public CellStyleBuilder DefaultWarning()
        {
            return new CellStyleBuilder(workbook)
                        .SetFont("Areal")
                        .SetFontColor(IndexedColors.Black)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.LightYellow);
        }

        public CellStyleBuilder DefaultError()
        {
            return new CellStyleBuilder(workbook)
                        .SetFont("Areal")
                        .SetFontColor(IndexedColors.Red)
                        .SetBold(false)
                        .SetBackground(FillPattern.SolidForeground, IndexedColors.DarkRed);
        }

        public CellStyleBuilder SetFont(string fontName)
        {
            font = workbook.CreateFont();
            Style.SetFont(font);
            return this;
        }

        public CellStyleBuilder SetFontColor(IndexedColors color)
        {
            font.Color = color.Index;
            return this;
        }        

        public CellStyleBuilder SetBold(bool isBold)
        {
            font.IsBold = isBold;
            return this;
        }

        public CellStyleBuilder SetBackground(FillPattern pattern, IndexedColors color)
        {
            Style.FillPattern = pattern;
            Style.FillForegroundColor = color.Index;
            return this;
        }
    }
}
