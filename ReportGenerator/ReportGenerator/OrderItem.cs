using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator
{
    /// <summary>
    /// One row of order
    /// </summary>
    class OrderItem
    {
        public int ProductCode { get; set; }
        public int Number { get; set; }
        public string Name { get; set; }
        public string Measurement { get; set; }
        public double Price { get; set; }
    }
}
