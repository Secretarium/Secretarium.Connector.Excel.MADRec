using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Xl = Microsoft.Office.Interop.Excel;

namespace Secretarium.Excel
{
    public static class ExcelHelpers
    {
        public static bool FindWorksheet(this Xl.Workbook wb, string name, out Xl.Worksheet ws)
        {
            ws = wb.Worksheets.OfType<Xl.Worksheet>().FirstOrDefault(x => x.Name == name);
            return ws != null;
        }

        public static string OffsetAddress(string address, int row, int col)
        {
            if (row == 0 && col == 0)
                return address;

            var regex = new Regex(@"([A-Z]+)(\d+)");
            var match = regex.Match(address);
            if (!match.Success)
                return address;
            
            int nCol = col;
            var ascii = Encoding.ASCII.GetBytes(match.Groups[1].Value);
            for (int i = ascii.Length - 1, j = 1; i >= 0; i--, j *= 26)
            {
                nCol += (ascii[i] - 64) * j;
            }

            int nRow = int.Parse(match.Groups[2].Value) + row;
            
            var columnName = "";
            while (nCol > 0)
            {
                int m = (nCol - 1) % 26;
                columnName = Convert.ToChar(65 + m).ToString() + columnName;
                nCol = (nCol - m) / 26;
            }

            return columnName + nRow;
        }
    }
}
