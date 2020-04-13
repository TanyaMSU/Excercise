using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ReportGenerator
{
    /// <summary>
    /// Order instance, collection of orderItem
    /// </summary>
    class Order : List<OrderItem>
    {
        public void FillOrder(int numOfItems, DataTable orderDataTable)
        {
            for (int j = 0; j < numOfItems; j++)
            {
                OrderItem newOrderItem = new OrderItem
                {
                    ProductCode = (int)orderDataTable.Rows[j]["Составляющие заказа.Код продукта"],
                    Number = (int)orderDataTable.Rows[j]["Количество"],
                    Name = Convert.ToString(orderDataTable.Rows[j]["Наименование "]),
                    Measurement = Convert.ToString(orderDataTable.Rows[j]["Единицы измерения"]),
                    Price = Convert.ToDouble(orderDataTable.Rows[j]["Цена"])
                };
                this.Add(newOrderItem);
            }
        }
    }
}
