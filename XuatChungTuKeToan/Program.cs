using System;
using System.Data;
using System.Linq;
using Aspose.Cells;

namespace XuatChungTuKeToan
{
    internal class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();

            Workbook wb = new Workbook();
            wb.Worksheets.RemoveAt(0); // Xóa sheet mặc định

            dt = PhieuNhapKhoDataForTesting();
            XuatPhieuNhapKho(wb, dt); // Xuất Phiếu nhập kho

            XuatPhieuThu(wb, 1200000); // Xuất Phiếu thu

            dt = GiayThanhToanTienTamUngDataForTesting();
            XuatGiayThanhToanTienTamUng(wb, dt);

            string filePath = @"ChungTuKeToan.xlsx";
            wb.Save(filePath);

            Console.WriteLine($"File location: " + System.IO.Path.GetFullPath(filePath));

            // Tự động mở file (Excel)
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = System.IO.Path.GetFullPath(filePath),
                UseShellExecute = true
            });
            Console.ReadLine();
        }
        #region "XuatPhieuNhapKho"
        static void XuatPhieuNhapKho(Workbook wb, DataTable dt)
        {
            Worksheet ws = wb.Worksheets.Add("PhieuNhapKho");

            // Khai báo đối tượng để sử dụng các thuộc tính style được định nghĩa
            var styles = new ExcelReportStyles(wb);

            // Phần mở đầu (thông tin chung)

            ws.Cells["A1"].PutValue("Đơn vị: ...");
            ws.MergeAndStyle("A1", "D1", styles.Bold);

            ws.Cells["A2"].PutValue("Bộ phận: ...");
            ws.MergeAndStyle("A2", "D2", styles.Bold);

            ws.Cells["E1"].PutValue("Mẫu số 01 - VT");
            ws.MergeAndStyle("E1", "H1", styles.CenterBold);

            ws.Cells["E2"].PutValue("(Kèm theo Thông tư số 99/2025/TT-BTC\nngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính)");
            ws.MergeAndStyle("E2", "H3", styles.CenterItalic);

            ws.Cells["A5"].PutValue("PHIẾU NHẬP KHO");
            ws.MergeAndStyle("A5", "H5", styles.Title);

            ws.Cells["A6"].PutValue("Ngày ..." + " tháng ..." + " năm ...");
            ws.MergeAndStyle("A6", "H6", styles.CenterItalic);

            ws.Cells["G7"].PutValue("Số: ...");
            ws.Cells.CreateRange("G7", "H7").Merge();

            ws.Cells["G8"].PutValue("Nợ: ...");
            ws.Cells.CreateRange("G8", "H8").Merge();

            ws.Cells["G9"].PutValue("Có: ...");
            ws.Cells.CreateRange("G9", "H9").Merge();

            ws.Cells["A10"].PutValue("- Họ và tên người giao: ...");
            ws.Cells.CreateRange("A10", "H10").Merge();

            ws.Cells["A11"].PutValue("- Theo ..." + " số ..." + " ngày ..." + " tháng ..." + " năm ..." + " của ...");
            ws.Cells.CreateRange("A11", "H11").Merge();

            ws.Cells["A12"].PutValue("- Nhập tại kho: ..." + " địa điểm ...");
            ws.Cells.CreateRange("A12", "H12").Merge();

            // Phần nội dung (chi tiết nội dung nghiệp vụ)

            // Tiêu đề bảng
            int titleRow = 13; // Dòng bắt đầu tiêu đề bảng (index: 13 - tương ứng dòng 14 trong Excel)

            ws.Cells["A14"].PutValue("STT");
            ws.MergeAndStyle("A14", "A15", styles.CenterBold);

            ws.Cells["B14"].PutValue("Tên, nhãn hiệu, quy cách, phẩm chất vật tư, dụng cụ sản phẩm, hàng hóa");
            ws.MergeAndStyle("B14", "B15", styles.CenterBold);

            ws.Cells["C14"].PutValue("Mã số");
            ws.MergeAndStyle("C14", "C15", styles.CenterBold);

            ws.Cells["D14"].PutValue("Đơn vị tính");
            ws.MergeAndStyle("D14", "D15", styles.CenterBold);

            ws.Cells["E14"].PutValue("Số lượng");
            ws.MergeAndStyle("E14", "F14", styles.CenterBold);

            ws.Cells["E15"].PutValue("Theo chứng từ");
            ws.Cells["E15"].SetStyle(styles.CenterBold);

            ws.Cells["F15"].PutValue("Thực nhập");
            ws.Cells["F15"].SetStyle(styles.CenterBold);

            ws.Cells["G14"].PutValue("Đơn giá");
            ws.MergeAndStyle("G14", "G15", styles.CenterBold);

            ws.Cells["H14"].PutValue("Thành tiền");
            ws.MergeAndStyle("H14", "H15", styles.CenterBold);

            // Dòng kí tự
            ws.Cells["A16"].PutValue("A");
            ws.Cells["B16"].PutValue("B");
            ws.Cells["C16"].PutValue("C");
            ws.Cells["D16"].PutValue("D");
            ws.Cells["E16"].PutValue("1");
            ws.Cells["F16"].PutValue("2");
            ws.Cells["G16"].PutValue("3");
            ws.Cells["H16"].PutValue("4");
            ws.Cells.CreateRange("A16", "H16").ApplyStyle(styles.Center, new StyleFlag { All = true });

            int startRow = 16; // Dòng bắt đầu đổ dữ liệu (index: 16 - tương ứng dòng 17 trong Excel)
            int startCol = 0; // Cột bắt đầu đổ dữ liệu (index: 0 - tương ứng cột A trong Excel)

            ImportTableOptions importDataOptions = new ImportTableOptions();
            importDataOptions.IsFieldNameShown = false; // Không hiện các thuộc tính của bảng
            importDataOptions.InsertRows = true; // Tự động thêm dòng

            // Thêm dữ liệu vào bảng
            ws.Cells.ImportData(dt, startRow, startCol, importDataOptions);

            int totalRows = dt.Rows.Count; // Tổng số dòng của DataTable
            int totalCols = dt.Columns.Count; // Tổng số cột của DataTable

            int currentRow = startRow + totalRows; // Dòng hiện tại (dòng bắt đầu đổ dữ liệu + Tổng số dòng của DataTable)

            // Áp dụng Style theo từng cột cho vùng dữ liệu từ dòng bắt đầu đổ dữ liệu (startRow)
            ws.SetColumnStyles(startRow, totalRows,
                // Mảng style truyền trực tiếp theo thứ tự cột:
                styles.Center,      // Cột A (STT)
                styles.Default,      // Cột B (Tên SP)
                styles.Center,      // Cột C (Mã số)
                styles.Center,      // Cột D (ĐVT)
                styles.Number2,     // Cột 1 (SLCT)
                styles.Number2,     // Cột 2 (SLTN)
                styles.Number3,     // Cột 3 (Đơn giá)
                styles.Number3      // Cột 4 (Thành tiền)
            );

            // Dòng tổng
            currentRow++;
            ws.Cells["B" + currentRow].PutValue("Cộng");
            ws.Cells["C" + currentRow].PutValue("x");
            ws.Cells["D" + currentRow].PutValue("x");
            ws.Cells["E" + currentRow].PutValue("x");
            ws.Cells["F" + currentRow].PutValue("x");
            ws.Cells["G" + currentRow].PutValue("x");
            ws.Cells.CreateRange("B" + currentRow, "G" + currentRow).ApplyStyle(styles.CenterBold, new StyleFlag { All = true });

            // Tính tổng cột Thành tiền
            decimal tongThanhTien = dt.Compute("SUM([ThanhTien])", "") is decimal val ? val : 0m;

            ws.Cells["H" + currentRow].PutValue(tongThanhTien);
            ws.Cells["H" + currentRow].SetStyle(styles.TongTien);

            // Tạo khung cho bảng
            // Xác định tổng số dòng của toàn bộ Bảng (Dòng hiện tại - Dòng bắt đầu tiêu đề bảng)
            int totalRowsTable = currentRow - titleRow;
            if (totalRowsTable > 0)
            {
                // Tạo Range cho Bảng
                Range tableRange = ws.Cells.CreateRange(titleRow, startCol, totalRowsTable, totalCols);
                // Kẻ khung theo Range
                tableRange.SetTableBorders();
            }

            currentRow += 2;
            ws.Cells["A" + currentRow].PutValue("- Tổng số tiền (viết bằng chữ): " + Libraries.NumberToVietnameseWords(tongThanhTien) + " đồng.");
            ws.Cells.CreateRange("A" + currentRow, "H" + currentRow).Merge();

            currentRow++;
            ws.Cells["A" + currentRow].PutValue("- Số chứng từ gốc kèm theo: ...");
            ws.Cells.CreateRange("A" + currentRow, "H" + currentRow).Merge();

            // Phần kết (xác thực và trách nhiệm pháp lý)

            currentRow += 2;
            ws.Cells["G" + currentRow].PutValue("Ngày ... tháng ... năm ...");
            ws.MergeAndStyle("G" + currentRow, "H" + currentRow, styles.CenterItalic);

            currentRow++;
            ws.Cells["A" + currentRow].PutValue("Người lập phiếu");
            ws.MergeAndStyle("A" + currentRow, "B" + currentRow, styles.CenterBold);

            ws.Cells["C" + currentRow].PutValue("Người giao hàng");
            ws.MergeAndStyle("C" + currentRow, "D" + currentRow, styles.CenterBold);

            ws.Cells["E" + currentRow].PutValue("Thủ kho");
            ws.MergeAndStyle("E" + currentRow, "F" + currentRow, styles.CenterBold);

            ws.Cells["G" + currentRow].PutValue("Kế toán trưởng");
            ws.MergeAndStyle("G" + currentRow, "H" + currentRow, styles.CenterBold);

            currentRow++;
            ws.Cells["A" + currentRow].PutValue("(Ký, họ tên)");
            ws.MergeAndStyle("A" + currentRow, "B" + currentRow, styles.CenterItalic);

            ws.Cells["C" + currentRow].PutValue("(Ký, họ tên)");
            ws.MergeAndStyle("C" + currentRow, "D" + currentRow, styles.CenterItalic);

            ws.Cells["E" + currentRow].PutValue("(Ký, họ tên)");
            ws.MergeAndStyle("E" + currentRow, "F" + currentRow, styles.CenterItalic);

            ws.Cells["G" + currentRow].PutValue("(Hoặc bộ phận");
            ws.MergeAndStyle("G" + currentRow, "H" + currentRow, styles.CenterItalic);

            currentRow++;
            ws.Cells["G" + currentRow].PutValue("có nhu cầu nhập)\n(Ký, họ tên)");
            ws.MergeAndStyle("G" + currentRow, "H" + currentRow, styles.CenterItalic);

            // Đặt độ rộng cột từ A -> H
            ws.SetColumnWidths(startCol, 8, 35, 12, 10, 12, 12, 15, 18);
            // Tự động căn chỉnh chiều cao dòng
            ws.AutoFitAllRows();

            // Page setup
            // Tiêu đề lặp lại khi sang trang (Dòng 14 -> 16 gồm Tiêu đề bảng + Dòng kí tự)
            ws.PageSetup.PrintTitleRows = "$14:$16";
            ws.PageSetup.PrintTitleColumns = "";
            // Khổ giấy & Hướng in
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
            ws.PageSetup.Orientation = PageOrientationType.Portrait; // Khổ dọc
            // Căn lề an toàn cho máy in (Đơn vị inch)
            ws.PageSetup.LeftMarginInch = 0.3;
            ws.PageSetup.RightMarginInch = 0.3;
            ws.PageSetup.TopMarginInch = 0.4;
            ws.PageSetup.BottomMarginInch = 0.5;
            ws.PageSetup.HeaderMarginInch = 0.2;
            ws.PageSetup.FooterMarginInch = 0.2;
            // Căn giữa trang in theo chiều ngang
            ws.PageSetup.CenterHorizontally = true;
            // Khóa vừa 1 trang chiều rộng, chiều dài cuộn tự do
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = 0;
            // Đánh số trang góc dưới bên phải (Ví dụ: Trang: 1/3)
            ws.PageSetup.SetFooter(2, "Trang: &P/&N");
        }
        #endregion
        #region "PhieuNhapKhoDataForTesting"
        private static DataTable PhieuNhapKhoDataForTesting()
        {
            DataTable dt = new DataTable("PhieuXuatKho");

            dt.Columns.Add("Stt", typeof(int));
            dt.Columns.Add("TenSanPham", typeof(string));
            dt.Columns.Add("MaSo", typeof(string));
            dt.Columns.Add("DonViTinh", typeof(string));
            dt.Columns.Add("SoLuongChungTu", typeof(decimal));
            dt.Columns.Add("SoLuongThucNhap", typeof(decimal));
            dt.Columns.Add("DonGia", typeof(decimal));
            dt.Columns.Add("ThanhTien", typeof(decimal));

            object[] sampleData = new object[]
            {
                new object[] {1, "Laptop Dell Inspiron 15", "SP001", "Cái", 5.0, 5.0, 15000000.0, 75000000.0},
                new object[] {2, "Chuột không dây Logitech M220", "SP002", "Cái", 20.0, 20.0, 250000.0, 5000000.0},
                new object[] {3, "Bàn phím cơ RGB Gaming", "SP003", "Cái", 15.0, 15.0, 850000.0, 12750000.0},
                new object[] {4, "Màn hình Dell 24 inch IPS", "SP004", "Cái", 10.0, 10.0, 3200000.0, 32000000.0},
                new object[] {5, "Tai nghe chụp tai Gaming", "SP005", "Cái", 12.0, 12.0, 600000.0, 7200000.0},
                new object[] {6, "Ổ cứng SSD Kingston 512GB", "SP006", "Cái", 25.0, 25.0, 950000.0, 23750000.0},
                new object[] {7, "RAM DDR4 8GB 3200MHz", "SP007", "Thanh", 30.0, 30.0, 600000.0, 18000000.0},
                new object[] {8, "Cáp HDMI 2.0 dài 1.8m", "SP008", "Sợi", 50.0, 50.0, 80000.0, 4000000.0},
                new object[] {9, "Hub USB Type-C 4 cổng", "SP009", "Cái", 15.0, 15.0, 220000.0, 3300000.0},
                new object[] {10, "Kệ kê laptop nhôm tản nhiệt", "SP010", "Cái", 10.0, 10.0, 180000.0, 1800000.0},
                new object[] {11, "Ghế xoay văn phòng lưng lưới", "SP011", "Cái", 8.0, 8.0, 1250000.0, 10000000.0},
                new object[] {12, "Bàn làm việc gỗ MDF 1m2", "SP012", "Cái", 5.0, 5.0, 950000.0, 4750000.0},
                new object[] {13, "Giấy A4 Double A 70gsm", "SP013", "Ram", 100.0, 100.0, 75000.0, 7500000.0},
                new object[] {14, "Bút bi Thiên Long 025", "SP014", "Hộp", 40.0, 40.0, 35000.0, 1400000.0},
                new object[] {15, "Bút dạ quang nhiều màu", "SP015", "Hộp", 15.0, 15.0, 45000.0, 675000.0},
                new object[] {16, "Sổ tay bìa da A5", "SP016", "Quyển", 50.0, 50.0, 25000.0, 1250000.0},
                new object[] {17, "Kẹp giấy bướm màu", "SP017", "Hộp", 30.0, 30.0, 12000.0, 360000.0},
                new object[] {18, "Băng keo trong 5cm", "SP018", "Cuộn", 60.0, 60.0, 15000.0, 900000.0},
                new object[] {19, "Kéo văn phòng lớn", "SP019", "Cái", 20.0, 20.0, 18000.0, 360000.0},
                new object[] {20, "Hộp đựng hồ sơ tài liệu", "SP020", "Cái", 25.0, 25.0, 30000.0, 750000.0},
                new object[] {21, "Máy in Laser Brother HL-L2321D", "SP021", "Cái", 3.0, 3.0, 2800000.0, 8400000.0},
                new object[] {22, "Mực in Brother TN-2385", "SP022", "Hộp", 10.0, 10.0, 450000.0, 4500000.0},
                new object[] {23, "Router Wifi Tenda AC10", "SP023", "Cái", 10.0, 10.0, 550000.0, 5500000.0},
                new object[] {24, "Dây mạng LAN Cat6", "SP024", "Mét", 200.0, 200.0, 7000.0, 1400000.0},
                new object[] {25, "Đầu nối cáp mạng RJ45", "SP025", "Cái", 30.0, 30.0, 10000.0, 300000.0},
                new object[] {26, "Loa vi tính 2.1", "SP026", "Bộ", 8.0, 8.0, 450000.0, 3600000.0},
                new object[] {27, "Webcam Full HD 1080p", "SP027", "Cái", 12.0, 12.0, 550000.0, 6600000.0},
                new object[] {28, "Đế tản nhiệt laptop 2 quạt", "SP028", "Cái", 15.0, 15.0, 150000.0, 2250000.0},
                new object[] {29, "Thẻ nhớ MicroSD 64GB", "SP029", "Cái", 25.0, 25.0, 120000.0, 3000000.0},
                new object[] {30, "USB 3.0 SanDisk 64GB", "SP030", "Cái", 20.0, 20.0, 140000.0, 2800000.0}
            };

            foreach (object[] row in sampleData)
            {
                dt.Rows.Add(row);
            }
            return dt;
        }
        #endregion
        #region "XuatPhieuThu"
        static void XuatPhieuThu(Workbook wb, decimal soTien)
        {
            Worksheet ws = wb.Worksheets.Add("PhieuThu");

            // Khai báo đối tượng để sử dụng các thuộc tính style được định nghĩa
            var styles = new ExcelReportStyles(wb);

            // Phần mở đầu (thông tin chung)

            ws.Cells["A1"].PutValue("Đơn vị: ...");
            ws.MergeAndStyle("A1", "B1", styles.Bold);

            ws.Cells["A2"].PutValue("Địa chỉ: ...");
            ws.MergeAndStyle("A2", "B3", styles.Bold);

            ws.Cells["C1"].PutValue("Mẫu số 01 - TT");
            ws.MergeAndStyle("C1", "E1", styles.CenterBold);

            ws.Cells["C2"].PutValue("(Kèm theo Thông tư số 99/2025/TT-BTC\nngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính)");
            ws.MergeAndStyle("C2", "E3", styles.CenterItalic);

            ws.Cells["A5"].PutValue("PHIẾU THU");
            ws.MergeAndStyle("A5", "E5", styles.Title);

            ws.Cells["A6"].PutValue("Ngày ..." + " tháng ..." + " năm ...");
            ws.MergeAndStyle("A6", "E6", styles.CenterItalic);

            ws.Cells["E7"].PutValue("Quyển số: ...");
            ws.Cells["E8"].PutValue("Số: ...");
            ws.Cells["E9"].PutValue("Nợ: ...");
            ws.Cells["E10"].PutValue("Có: ...");

            ws.Cells["A12"].PutValue("- Họ và tên người nộp tiền: ...");
            ws.Cells.CreateRange("A12", "E12").Merge();

            ws.Cells["A14"].PutValue("- Địa chỉ: ...");
            ws.Cells.CreateRange("A14", "E14").Merge();

            ws.Cells["A16"].PutValue("- Lý do nộp: ...");
            ws.Cells.CreateRange("A16", "E16").Merge();

            ws.Cells["A18"].PutValue("- Số tiền: " + $"{soTien:N0}đ" + " (Viết bằng chữ): " + Libraries.NumberToVietnameseWords(soTien) + " đồng.");
            ws.Cells.CreateRange("A18", "E18").Merge();

            ws.Cells["A20"].PutValue("- Kèm theo: ..." + " Chứng từ gốc: ...");
            ws.Cells.CreateRange("A20", "E20").Merge();

            // Phần kết (xác thực và trách nhiệm pháp lý)

            ws.Cells["D22"].PutValue("Ngày ... tháng ... năm ...");
            ws.MergeAndStyle("D22", "E22", styles.CenterItalic);

            ws.Cells["A24"].PutValue("Giám đốc");
            ws.Cells["A24"].SetStyle(styles.CenterBold);

            ws.Cells["B24"].PutValue("Kế toán trưởng");
            ws.Cells["B24"].SetStyle(styles.CenterBold);

            ws.Cells["C24"].PutValue("Người nộp tiền");
            ws.Cells["C24"].SetStyle(styles.CenterBold);

            ws.Cells["D24"].PutValue("Người lập phiếu");
            ws.Cells["D24"].SetStyle(styles.CenterBold);

            ws.Cells["E24"].PutValue("Thủ quỹ");
            ws.Cells["E24"].SetStyle(styles.CenterBold);

            ws.Cells["A25"].PutValue("(Ký, họ tên,\nđóng dấu)");
            ws.MergeAndStyle("A25", "A26", styles.CenterItalic);

            ws.Cells["B25"].PutValue("(Ký, họ tên)");
            ws.Cells["B25"].SetStyle(styles.CenterItalic);

            ws.Cells["C25"].PutValue("(Ký, họ tên)");
            ws.Cells["C25"].SetStyle(styles.CenterItalic);

            ws.Cells["D25"].PutValue("(Ký, họ tên)");
            ws.Cells["D25"].SetStyle(styles.CenterItalic);

            ws.Cells["E25"].PutValue("(Ký, họ tên)");
            ws.Cells["E25"].SetStyle(styles.CenterItalic);

            // Đặt độ rộng cột từ A -> H
            ws.SetColumnWidths(0, 20, 20, 20, 20, 20);
            // Tự động căn chỉnh chiều cao dòng
            ws.AutoFitAllRows();

            // Page setup
            // Tiêu đề lặp lại khi sang trang (Dòng 14 -> 16 gồm Tiêu đề bảng + Dòng kí tự)
            ws.PageSetup.PrintTitleRows = "";
            ws.PageSetup.PrintTitleColumns = "";
            // Khổ giấy & Hướng in
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
            ws.PageSetup.Orientation = PageOrientationType.Portrait; // Khổ dọc
            // Căn lề an toàn cho máy in (Đơn vị inch)
            ws.PageSetup.LeftMarginInch = 0.3;
            ws.PageSetup.RightMarginInch = 0.3;
            ws.PageSetup.TopMarginInch = 0.4;
            ws.PageSetup.BottomMarginInch = 0.5;
            ws.PageSetup.HeaderMarginInch = 0.2;
            ws.PageSetup.FooterMarginInch = 0.2;
            // Căn giữa trang in theo chiều ngang
            ws.PageSetup.CenterHorizontally = true;
            // Khóa vừa 1 trang chiều rộng, chiều dài cuộn tự do
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = 0;
            // Đánh số trang góc dưới bên phải (Ví dụ: Trang: 1/3)
            ws.PageSetup.SetFooter(2, "Trang: &P/&N");
        }
        #endregion
        #region "XuatGiayThanhToanTienTamUng"
        static void XuatGiayThanhToanTienTamUng(Workbook wb, DataTable dt)
        {
            Worksheet ws = wb.Worksheets.Add("GiayThanhToanTienTamUng");

            // Khai báo đối tượng để sử dụng các thuộc tính style được định nghĩa
            var styles = new ExcelReportStyles(wb);

            // Phần mở đầu (thông tin chung)

            ws.Cells["A1"].PutValue("Đơn vị: ...");
            ws.MergeAndStyle("A1", "B1", styles.Bold);

            ws.Cells["A2"].PutValue("Bộ phận: ...");
            ws.MergeAndStyle("A2", "B2", styles.Bold);

            ws.Cells["C1"].PutValue("Mẫu số 04 - TT");
            ws.MergeAndStyle("C1", "D1", styles.CenterBold);

            ws.Cells["C2"].PutValue("(Kèm theo Thông tư số 99/2025/TT-BTC\nngày 27 tháng 10 năm 2025 của Bộ trưởng Bộ Tài chính)");
            ws.MergeAndStyle("C2", "D3", styles.CenterItalic);

            ws.Cells["A5"].PutValue("GIẤY THANH TOÁN TIỀN TẠM ỨNG");
            ws.MergeAndStyle("A5", "D5", styles.Title);

            ws.Cells["A6"].PutValue("Ngày ..." + " tháng ..." + " năm ...");
            ws.MergeAndStyle("A6", "D6", styles.CenterItalic);

            ws.Cells["D7"].PutValue("Số: ...");
            ws.Cells["D8"].PutValue("Nợ: ...");
            ws.Cells["D9"].PutValue("Có: ...");

            ws.Cells["A10"].PutValue("- Họ và tên người thanh toán: ...");
            ws.Cells.CreateRange("A10", "D10").Merge();

            ws.Cells["A11"].PutValue("- Bộ phận (hoặc địa chỉ): ...");
            ws.Cells.CreateRange("A11", "D11").Merge();

            ws.Cells["A12"].PutValue("- Số tiền tạm ứng được thanh toán theo bảng dưới đây:");
            ws.Cells.CreateRange("A12", "D12").Merge();

            // Phần nội dung (chi tiết nội dung nghiệp vụ)

            // Tiêu đề bảng
            int titleRow = 12; // Dòng bắt đầu tiêu đề bảng (index: 12 - tương ứng dòng 13 trong Excel)

            ws.Cells["A13"].PutValue("Diễn giải");
            ws.MergeAndStyle("A13", "C13", styles.CenterBold);

            ws.Cells["D13"].PutValue("Số tiền");
            ws.Cells["D13"].SetStyle(styles.CenterBold);

            // Dòng kí tự
            ws.Cells["A14"].PutValue("A");
            ws.MergeAndStyle("A14", "C14", styles.Center);

            ws.Cells["D14"].PutValue("1");
            ws.Cells["D14"].SetStyle(styles.Center);

            int startRow = 14; // Dòng bắt đầu đổ dữ liệu (index: 14 - tương ứng dòng 15 trong Excel)
            int startCol = 0; // Cột bắt đầu đổ dữ liệu (index: 0 - tương ứng cột A trong Excel)
            int totalRows = dt.Rows.Count; // Tổng số dòng của DataTable
            int totalCols = dt.Columns.Count; // Tổng số cột của DataTable
            int currentRow = startRow; // Dòng hiện tại (dòng bắt đầu đổ dữ liệu)

            if (totalRows > 0)
            {
                // Chèn trước số lượng dòng dữ liệu chuẩn bị ghi
                ws.Cells.InsertRows(startRow, totalRows);

                for(int i = 0; i < totalRows; i++)
                {
                    // Lấy từng dòng dữ liệu trong DataTable xử lý
                    DataRow row = dt.Rows[i];

                    // Ghi Diễn giải vào Cột A & Gộp A -> C
                    ws.Cells[currentRow, 0].PutValue(row[0]);
                    ws.Cells.CreateRange(currentRow, 0, 1, 3).Merge();
                    // Ghi Số tiền vào Cột D
                    if (row[1] != DBNull.Value && row[1] != null)
                    {
                        ws.Cells[currentRow, 3].PutValue(row[1]);
                    }
                    currentRow++; // Xuống dòng
                }
            }

            // Áp dụng Style theo từng cột cho vùng dữ liệu từ dòng bắt đầu đổ dữ liệu (startRow)
            ws.SetColumnStyles( startRow, totalRows,
                // Mảng style truyền trực tiếp theo thứ tự cột:
                styles.Default,           // Cột A (Diễn giải)
                null,
                null,           
                styles.Number3  // Cột 1 (Số tiền)
            );

            // Tạo khung cho bảng
            // Xác định tổng số dòng của toàn bộ Bảng (Dòng hiện tại - Dòng bắt đầu tiêu đề bảng)
            int totalRowsTable = currentRow - titleRow;
            if (totalRowsTable > 0)
            {
                // Tạo Range cho Bảng
                Range tableRange = ws.Cells.CreateRange(titleRow, startCol, totalRowsTable, totalCols + 2);
                // Kẻ khung theo Range
                tableRange.SetTableBorders();
            }

            // Phần kết (xác thực và trách nhiệm pháp lý)

            currentRow += 2;
            ws.Cells["A" + currentRow].PutValue("Giám đốc");
            ws.Cells["A" + currentRow].SetStyle(styles.CenterBold);

            ws.Cells["B" + currentRow].PutValue("Kế toán trưởng");
            ws.Cells["B" + currentRow].SetStyle(styles.CenterBold);

            ws.Cells["C" + currentRow].PutValue("Kế toán thanh toán");
            ws.Cells["C" + currentRow].SetStyle(styles.CenterBold);

            ws.Cells["D" + currentRow].PutValue("Người đề nghị");
            ws.Cells["D" + currentRow].SetStyle(styles.CenterBold);

            currentRow++;
            ws.Cells["A" + currentRow].PutValue("(Ký, họ tên)");
            ws.Cells["A" + currentRow].SetStyle(styles.CenterItalic);

            ws.Cells["B" + currentRow].PutValue("(Ký, họ tên)");
            ws.Cells["B" + currentRow].SetStyle(styles.CenterItalic);

            ws.Cells["C" + currentRow].PutValue("(Ký, họ tên)");
            ws.Cells["C" + currentRow].SetStyle(styles.CenterItalic);

            ws.Cells["D" + currentRow].PutValue("thanh toán");
            ws.Cells["D" + currentRow].SetStyle(styles.CenterBold);

            currentRow++;
            ws.Cells["D" + currentRow].PutValue("(Ký, họ tên)");
            ws.Cells["D" + currentRow].SetStyle(styles.CenterItalic);

            // Đặt độ rộng cột từ A -> H
            ws.SetColumnWidths(startCol, 20, 20, 25, 30);
            // Tự động căn chỉnh chiều cao dòng
            ws.AutoFitAllRows();

            // Page setup
            // Tiêu đề lặp lại khi sang trang (Dòng 14 -> 16 gồm Tiêu đề bảng + Dòng kí tự)
            ws.PageSetup.PrintTitleRows = "$13:$14";
            ws.PageSetup.PrintTitleColumns = "";
            // Khổ giấy & Hướng in
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;
            ws.PageSetup.Orientation = PageOrientationType.Portrait; // Khổ dọc
            // Căn lề an toàn cho máy in (Đơn vị inch)
            ws.PageSetup.LeftMarginInch = 0.3;
            ws.PageSetup.RightMarginInch = 0.3;
            ws.PageSetup.TopMarginInch = 0.4;
            ws.PageSetup.BottomMarginInch = 0.5;
            ws.PageSetup.HeaderMarginInch = 0.2;
            ws.PageSetup.FooterMarginInch = 0.2;
            // Căn giữa trang in theo chiều ngang
            ws.PageSetup.CenterHorizontally = true;
            // Khóa vừa 1 trang chiều rộng, chiều dài cuộn tự do
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = 0;
            // Đánh số trang góc dưới bên phải (Ví dụ: Trang: 1/3)
            ws.PageSetup.SetFooter(2, "Trang: &P/&N");
        }
        #endregion
        #region "GiayThanhToanTienTamUngDataForTesting"
        public static DataTable GiayThanhToanTienTamUngDataForTesting()
        {
            DataTable dt = new DataTable("GiayThanhToanTamUng");
            dt.Columns.Add("DienGiai", typeof(string));
            dt.Columns.Add("SoTien", typeof(object)); // Dùng object để hỗ trợ nạp decimal hoặc DBNull.Value

            // ==========================================
            // 1. KHỞI TẠO DỮ LIỆU ĐỘNG (Từ CSDL / API)
            // ==========================================
            decimal soTamUngKyTruoc = 2500000m;

            // Danh sách N Phiếu chi tạm ứng (ĐỘNG)
            var dsPhieuChiTamUng = new[]
            {
                new { SoPhieu = "PC001/10", Ngay = "05/10/2026", SoTien = 5000000m, NoiDung = "Tạm ứng đợt 1 đi công tác" },
                new { SoPhieu = "PC008/10", Ngay = "12/10/2026", SoTien = 3000000m, NoiDung = "Tạm ứng đợt 2 mua vé máy bay" },
                new { SoPhieu = "PC015/10", Ngay = "18/10/2026", SoTien = 1500000m, NoiDung = "Tạm ứng bổ sung chi phí xe bãi" }
            };

            // Danh sách M Chứng từ chi thực tế (ĐỘNG)
            var dsChungTuChi = new[]
            {
                new { SoChungTu = "HD0123", Ngay = "15/10/2026", SoTien = 4200000m, NoiDung = "Vé máy bay khứ hồi Hà Nội - TP.HCM" },
                new { SoChungTu = "HD0456", Ngay = "18/10/2026", SoTien = 3500000m, NoiDung = "Tiền phòng khách sạn 3 đêm" },
                new { SoChungTu = "HD0789", Ngay = "19/10/2026", SoTien = 1800000m, NoiDung = "Tiền tiếp khách đối tác" },
                new { SoChungTu = "BL0012", Ngay = "20/10/2026", SoTien = 650000m, NoiDung = "Cước Taxi di chuyển công tác" },
                new { SoChungTu = "HD0999", Ngay = "21/10/2026", SoTien = 450000m, NoiDung = "In ấn tài liệu hội thảo" }
            };

            // ==========================================
            // 2. TÍNH TOÁN NGHIỆP VỤ TỔNG CỘNG
            // ==========================================
            decimal tongPhieuChiKyNay = dsPhieuChiTamUng.Sum(x => x.SoTien);
            decimal tongMucI = soTamUngKyTruoc + tongPhieuChiKyNay; // Total I = 12,000,000

            decimal tongMucII = dsChungTuChi.Sum(x => x.SoTien); // Total II = 10,600,000

            decimal chenhLech = tongMucI - tongMucII; // 1,400,000 (Dương -> Chi không hết)
            decimal tamUngChiKhongHet = chenhLech > 0 ? chenhLech : 0m;
            decimal chiQuaTamUng = chenhLech < 0 ? Math.Abs(chenhLech) : 0m;

            // ==========================================
            // 3. ĐỔ DỮ LIỆU VÀO DATATABLE THEO TUYẾN TÍNH
            // ==========================================

            // --- MỤC I: SỐ TIỀN TẠM ỨNG ---
            dt.Rows.Add("I. Số tiền tạm ứng", tongMucI);
            dt.Rows.Add("   1. Số tạm ứng các kỳ trước chưa chi hết", soTamUngKyTruoc);
            dt.Rows.Add("   2. Số tạm ứng kỳ này:", DBNull.Value);

            // Vòng lặp đổ danh sách N phiếu chi tạm ứng
            foreach (var pc in dsPhieuChiTamUng)
            {
                dt.Rows.Add($"      - Phiếu chi số {pc.SoPhieu} ngày {pc.Ngay}", pc.SoTien);
            }

            // --- MỤC II: SỐ TIỀN ĐÃ CHI ---
            dt.Rows.Add("II. Số tiền đã chi", tongMucII);

            // Vòng lặp đổ danh sách M chứng từ chi thực tế
            int stt = 1;
            foreach (var ct in dsChungTuChi)
            {
                dt.Rows.Add($"   {stt}. Chứng từ số {ct.SoChungTu} ngày {ct.Ngay} ({ct.NoiDung})", ct.SoTien);
                stt++;
            }

            // --- MỤC III: CHÊNH LỆCH ---
            dt.Rows.Add("III. Chênh lệch", DBNull.Value);
            dt.Rows.Add("   1. Số tạm ứng chi không hết ( I - II )", tamUngChiKhongHet > 0 ? tamUngChiKhongHet : (object)DBNull.Value);
            dt.Rows.Add("   2. Chi quá số tạm ứng ( II - I )", chiQuaTamUng > 0 ? chiQuaTamUng : (object)DBNull.Value);

            return dt;
        }
        #endregion
    }
}
    