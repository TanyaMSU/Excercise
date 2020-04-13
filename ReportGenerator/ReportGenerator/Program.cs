using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace ReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(@"Please enter the path to containing folder like C:\Users\Tanya :");
            string pathToContainingfFolder = Console.ReadLine();

            //Declare some variables
            AccessDBReader accessDBReader;
            DataTable orderDataTable;
            int order;
            int numOfItems;

            //Get the table of all orders from DB and the number of orders
            accessDBReader = new AccessDBReader(pathToContainingfFolder);
            DataTable ordersDataTable = accessDBReader.GetOrdersDataTable();
            int numOfOrders = ordersDataTable.Rows.Count;

            //For each order
            for (int i = 0; i < numOfOrders; i++)
            {
                order = i + 1;
                //Get the table with all order info and the number of its items
                accessDBReader = new AccessDBReader(pathToContainingfFolder);
                orderDataTable = accessDBReader.GetOrderDataTable(order);
                numOfItems = orderDataTable.Rows.Count;

                //Declare and initialize order instance
                Order newOrder = new Order();
                newOrder.FillOrder(numOfItems, orderDataTable);

                //Get the table with some info about current order and shop number from it
                accessDBReader = new AccessDBReader(pathToContainingfFolder);
                DataTable orderRowDataTable = accessDBReader.GetOrderRowDataTable(order);
                int shopNumber = (int)orderRowDataTable.Rows[0]["№ магазина"];

                //Get the table with all shop info 
                accessDBReader = new AccessDBReader(pathToContainingfFolder);
                DataTable shopInfoDataTable = accessDBReader.GetShopInfoDataTable(shopNumber);

                //Create report only if shop info exists in database
                if (shopInfoDataTable.Rows.Count != 0)
                {
                    Shop shop = new Shop(shopNumber, shopInfoDataTable);

                    //Create report only if the date of order is present and the draft is absent
                    if ((orderRowDataTable.Rows[0]["Дата согласования закааза"] != null) && (Convert.ToString(orderRowDataTable.Rows[0]["Черновик заказа"]) == ""))
                    {
                        Report newReport = new Report(order, newOrder, shop);
                        newReport.GenerateReportForOneOrder(order, pathToContainingfFolder);

                    }
                    else
                    {
                        Console.WriteLine("Order date is absent or draft is present for order {0}", order);
                    }
                }
                else
                {
                    Console.WriteLine("There is no information about shop {0}", shopNumber);
                }
            }
            Console.WriteLine("Work completed!");
            Console.WriteLine("Press any key to close console");
            Console.ReadKey();
        }
    }
}
