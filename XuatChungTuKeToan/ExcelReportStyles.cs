using Aspose.Cells;

namespace XuatChungTuKeToan
{
    public class ExcelReportStyles
    {
        public Style Center { get; private set; }
        public Style Bold { get; private set; }
        public Style CenterBold { get; private set; }
        public Style CenterItalic { get; private set; }
        public Style Title { get; private set; }
        public Style Number2 { get; private set; }
        public Style Number3 { get; private set; }
        public Style TongTien { get; private set; }

        public ExcelReportStyles(Workbook wb)
        {
            // Áp dụng style mặc định cho toàn bộ workbook
            Style defaultStyle = wb.DefaultStyle;
            defaultStyle.Font.Name = "Times New Roman";
            defaultStyle.Font.Size = 11;
            defaultStyle.HorizontalAlignment = TextAlignmentType.Left;
            defaultStyle.VerticalAlignment = TextAlignmentType.Center;
            defaultStyle.IsTextWrapped = true;

            wb.DefaultStyle = defaultStyle;

            // Khai báo style con (kế thừa tự động từ DefaultStyle)
            Center = wb.CreateStyle();
            Center.HorizontalAlignment = TextAlignmentType.Center;

            Bold = wb.CreateStyle();
            Bold.Font.IsBold = true;

            CenterBold = wb.CreateStyle();
            CenterBold.Copy(Center);
            CenterBold.Font.IsBold = true;

            CenterItalic = wb.CreateStyle();
            CenterItalic.Copy(Center);
            CenterItalic.Font.IsItalic = true;

            Title = wb.CreateStyle();
            Title.Copy(CenterBold);
            Title.Font.Size = 13;

            Number2 = wb.CreateStyle();
            Number2.Number = 2;
            Number2.HorizontalAlignment = TextAlignmentType.Right;

            Number3 = wb.CreateStyle();
            Number3.Number = 3;
            Number3.HorizontalAlignment = TextAlignmentType.Right;

            TongTien = wb.CreateStyle();
            TongTien.Copy(CenterBold);
            TongTien.Number = 3;
        }
    }
}
