using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlToDb
{
    public class Function
    {
        public static DateTime ExcelToDate(string serial)
        {
            if (serial == null)
            {
                throw new ArgumentNullException(nameof(serial));
            }

            int serialDate = int.Parse(serial);
            if (serialDate > 59) serialDate--;
            return new DateTime(1899, 12, 31).AddDays(serialDate);
        }
    }
}
