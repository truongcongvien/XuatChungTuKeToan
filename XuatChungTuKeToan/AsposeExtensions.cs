using Aspose.Cells;

namespace XuatChungTuKeToan
{
    public static class AsposeExtensions
    {
        // Gộp ô cho toàn bộ các ô trong Range và áp dụng style cho toàn bộ Range
        public static void MergeAndStyle(this Worksheet ws, string upperLeftCell, string lowerRightCell, Style style)
        {
            Range range = ws.Cells.CreateRange(upperLeftCell, lowerRightCell);
            range.Merge();
            range.ApplyStyle(style, new StyleFlag { All = true });
        }
        //
        public static void SetColumnStyles(this Worksheet ws, int startRow, int totalRows, params Style[] styles)
        {
            if (styles == null || styles.Length == 0 || totalRows <= 0)
            {
                return;
            }
            StyleFlag flag = new StyleFlag { All = true };
            for (int colIndex = 0; colIndex < styles.Length; colIndex++)
            {
                if (styles[colIndex] == null)
                {
                    continue;
                }
                // Tạo Range cho toàn bộ cột từ startRow đến hết dữ liệu (Chỉ rộng 1 cột)
                Range range = ws.Cells.CreateRange(startRow, colIndex, totalRows, 1);
                range.ApplyStyle(styles[colIndex], flag);
            }
        }
        // Kẻ đường viền mảnh (Thin) cho toàn bộ các ô trong Range và kẻ viền ngoài (Outer Border)
        public static void SetTableBorders(this Range range, CellBorderType borderType = CellBorderType.Thin)
        {
            Workbook wb = range.Worksheet.Workbook;
            // Tạo Style kẻ viền toàn bộ (Bốn cạnh)
            Style borderStyle = wb.CreateStyle();
            borderStyle.Borders[BorderType.TopBorder].LineStyle = borderType;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = borderType;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = borderType;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = borderType;
            // Tạo StyleFlag chỉ bật thuộc tính Borders (giữ nguyên Font, Alignment... của các ô bên trong)
            StyleFlag borderFlag = new StyleFlag
            {
                Borders = true
            };
            // Phủ viền hàng loạt cho cả vùng
            range.ApplyStyle(borderStyle, borderFlag);
        }
        // Thiết lập độ rộng hàng loạt cho các cột bắt đầu từ startCol
        public static void SetColumnWidths(this Worksheet ws, int startCol, params double[] widths)
        {
            for (int i = 0; i < widths.Length; i++)
            {
                ws.Cells.SetColumnWidth(startCol + i, widths[i]);
            }
        }
        // Tự động căn chỉnh chiều cao dòng cho toàn bộ Sheet (Hỗ trợ cả ô gộp)
        public static void AutoFitAllRows(this Worksheet ws)
        {
            AutoFitterOptions options = new AutoFitterOptions();
            options.AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine;
            ws.AutoFitRows(options);
        }
    }
}
