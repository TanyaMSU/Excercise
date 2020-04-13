using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ReportGenerator
{
    /// <summary>
    /// Shop instance
    /// </summary>
    class Shop
    {
        public int Number { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Director { get; set; }

        public Shop(int shopNumber, DataTable shopInfoDataTable)
        {
            Number = shopNumber;
            Name = Convert.ToString(shopInfoDataTable.Rows[0]["Название магазина"]);
            Address = Convert.ToString(shopInfoDataTable.Rows[0]["Адрес магазина"]);
            Director = Convert.ToString(shopInfoDataTable.Rows[0]["Директор магазина"]);
        }
    }
}
