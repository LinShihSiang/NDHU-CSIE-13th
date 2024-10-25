using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NDHU_CSIE_13th.Extensions
{
    public static class XLWorkbookExtensions
    {
        public static void SetHeaders(this IXLWorksheet worksheet, IEnumerable<string> headers)
        {
            for (int i = 0; i < headers.Count(); i++)
            {
                worksheet.Cell(1, i + 1).Value = headers.ElementAt(i);
                worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(252, 229, 205);
            }
        }
    }
}
