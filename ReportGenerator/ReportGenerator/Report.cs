using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ReportGenerator
{
    /// <summary>
    /// Creates Excel file with report
    /// </summary>
    class Report
    {
        Shop shop;
        Order order;
        int orderNumber;

        public Report(int orderNum, Order givenOrder, Shop givenShop)
        {
            shop = givenShop;
            order = givenOrder;
            orderNumber = orderNum;
            
        }

        public void GenerateReportForOneOrder(int orderNumber, string pathToContainingFolder)
        {
            string fileTarget = pathToContainingFolder.Replace("/",@"\") + @"\Черновик" + orderNumber + ".xlsx";
            string fileTemplate = pathToContainingFolder.Replace("/", @"\") + @"\шаблон.xlsx";

            Application excelApp = new Application();
            Workbook wbTemp;
            Worksheet sh;

            wbTemp = excelApp.Workbooks.Open(fileTemplate);

            sh = wbTemp.Worksheets["Лист2"];

            sh.Cells[2, 2] = "Заказ № " + orderNumber;
            sh.Cells[4, 2] = "Магазин № " + shop.Number + " по адресу " + shop.Address;
            sh.Cells[12, 3] = "Директор " + shop.Director;

            Range range = (Range)sh.Rows[7];
            for (int i = 0; i < order.Count - 1; i++)
            {
                range.Insert(XlDirection.xlDown);
            }

            Range range2 = (Range)sh.Rows[7 + order.Count - 1];

            for (int i = 0; i < order.Count - 1; i++)
            {
                range2.Copy(sh.Rows[7 + i]);
            }

            for (int i = 0; i < order.Count; i++)
            {
                sh.Cells[7 + i, 1] = i+1;
                sh.Cells[7 + i, 2] = order[i].Name;
                sh.Cells[7 + i, 3] = order[i].Measurement;
                sh.Cells[7 + i, 4] = order[i].Price;
                sh.Cells[7 + i, 5] = order[i].Number;
                sh.Cells[7 + i, 6] = String.Format("=D{0}*E{1}", 7 + i, 7 + i);
            }

            sh.Cells[7 + order.Count, 6] = String.Format("=SUM(F{0}:F{1})", 7, 7 + order.Count - 1);
            sh.Cells[8 + order.Count, 6] = String.Format("=F{0}*0.2",7 + order.Count);
            sh.Cells[9 + order.Count, 6] = String.Format("=F{0}+F{1}", 7 + order.Count, 8 + order.Count);

            wbTemp.SaveAs(fileTarget);

            wbTemp.Close();
            Console.WriteLine("Report about order {0} is generated", orderNumber);
        }
    }
}
