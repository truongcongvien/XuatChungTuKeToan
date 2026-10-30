using System;
using System.Text;
using System.Text.RegularExpressions;

namespace XuatChungTuKeToan
{
    public static class Libraries
    {
        private static readonly string[] ChuSo = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

        // Đọc số kiểu decimal thành chữ theo chuẩn tiếng Việt, viết hoa chữ cái đầu
        public static string NumberToVietnameseWords(this decimal number)
        {
            // Lấy phần nguyên để xử lý
            decimal absNumber = Math.Floor(Math.Abs(number));

            // Xử lý riêng trường hợp số 0
            if (absNumber == 0)
                return "Không";

            string result = "";
            int billionLevel = 0; // Đếm số cấp "tỷ" (0: <1 tỷ, 1: tỷ, 2: tỷ tỷ,...)

            // Tách từng nhóm 9 chữ số (mỗi nhóm tương ứng 1 cấp Tỷ)
            do
            {
                decimal currentTyGroup = absNumber % 1_000_000_000m;
                absNumber = Math.Floor(absNumber / 1_000_000_000m);

                if (currentTyGroup > 0)
                {
                    string groupText = ReadGroupUnderBillion(currentTyGroup, absNumber > 0);

                    // Ghép đuôi "tỷ" tương ứng với cấp độ (tỷ, tỷ tỷ,...)
                    string tySuffix = "";
                    for (int i = 0; i < billionLevel; i++)
                    {
                        tySuffix += " tỷ";
                    }

                    result = $"{groupText}{tySuffix} {result}".Trim();
                }

                billionLevel++;
            } while (absNumber > 0);

            // Chuẩn hóa khoảng trắng dư thừa
            result = Regex.Replace(result.Trim(), @"\s+", " ");

            // Thêm từ "Âm" nếu là số âm
            if (number < 0)
            {
                result = "âm " + result;
            }

            // Viết hoa chữ cái đầu tiên
            return char.ToUpper(result[0]) + result.Substring(1);
        }

        // Đọc nhóm dưới 1 tỷ (tối đa 9 chữ số: gồm các đơn vị Trăm, Nghìn, Triệu)
        private static string ReadGroupUnderBillion(decimal number, bool hasHigherBlock)
        {
            string[] units = { "", "nghìn", "triệu" };
            string result = "";
            int blockIndex = 0;

            do
            {
                int triple = (int)(number % 1000m);
                number = Math.Floor(number / 1000m);

                int ones = triple % 10;
                int tens = (triple / 10) % 10;
                int hundreds = triple / 100;

                if (hundreds > 0 || tens > 0 || ones > 0)
                {
                    bool hasHigher = number > 0 || hasHigherBlock;
                    string blockText = ReadTriple(hundreds, tens, ones, hasHigher);
                    string unit = units[blockIndex];

                    result = $"{blockText} {unit} {result}".Trim();
                }

                blockIndex++;
            } while (number > 0);

            return result;
        }

        // Đọc nhóm 3 chữ số (trăm, chục, đơn vị)
        private static string ReadTriple(int hundreds, int tens, int ones, bool hasHigherBlock)
        {
            StringBuilder sb = new StringBuilder();

            // Hàng trăm
            if (hundreds > 0 || hasHigherBlock)
            {
                sb.Append(ChuSo[hundreds] + " trăm ");
            }

            // Hàng chục
            if (tens > 1)
            {
                sb.Append(ChuSo[tens] + " mươi ");
            }
            else if (tens == 1)
            {
                sb.Append("mười ");
            }
            else if (tens == 0 && ones > 0 && (hundreds > 0 || hasHigherBlock))
            {
                sb.Append("lẻ ");
            }

            // Hàng đơn vị
            if (ones > 0)
            {
                if (ones == 1 && tens > 1)
                {
                    sb.Append("mốt");
                }
                else if (ones == 4 && tens > 1)
                {
                    sb.Append("tư");
                }
                else if (ones == 5 && tens > 0)
                {
                    sb.Append("lăm");
                }
                else
                {
                    sb.Append(ChuSo[ones]);
                }
            }

            return sb.ToString().Trim();
        }
    }
}
