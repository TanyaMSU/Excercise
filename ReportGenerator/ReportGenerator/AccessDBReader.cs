using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace ReportGenerator
{
    /// <summary>
    /// Get tables from Access DataBase
    /// </summary>
    public class AccessDBReader
    {
        public readonly string strSQL;
        string connectionStr;
        string strSQLOrder = "SELECT [Составляющие заказа].*, [Библиотека продуктов].* " +
            "FROM(([Заказ на доставку] " +
            "INNER JOIN[Составляющие заказа] ON[Заказ на доставку].[№ заказа] =[Составляющие заказа].[№ заказа]) " +
            "INNER JOIN[Библиотека продуктов] ON[Библиотека продуктов].[Код продукта] =[Составляющие заказа].[Код продукта])WHERE[Заказ на доставку].[№ заказа] = ";
        string strSQLAllOrders = "SELECT [№ заказа] FROM [Заказ на доставку]";
        string strSQLOrderInfo = "SELECT [№ магазина], [Дата согласования закааза], [Черновик заказа] FROM [Заказ на доставку] WHERE [№ заказа] = ";
        string strSQLShop = "SELECT [Название магазина], [Адрес магазина], [Директор магазина] FROM [Магазин] WHERE [№ магазина] = ";

        public AccessDBReader(string pathToContainingFolder)
        {
            connectionStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathToContainingFolder.Replace("/", @"\") + @"\База данных.accdb;";
        }

        public DataTable GetDataFromAccessDB(string connectionString, string strSQL)
        {
            DataTable table = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Create a command and set its connection  
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                // Open the connection and execute the select command.  
                try
                {
                    // Open connecton  
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = command;
                    adapter.Fill(table);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                // The connection is automatically closed becasuse of using block.  
            }
            return table;
        }

        public DataTable GetOrdersDataTable()
        {
            return GetDataFromAccessDB(connectionStr, strSQLAllOrders);
        }

        public DataTable GetOrderDataTable(int order)
        {
            return GetDataFromAccessDB(connectionStr, strSQLOrder + Convert.ToString(order) + ";");
        }

        public DataTable GetOrderRowDataTable(int order)
        {
            return GetDataFromAccessDB(connectionStr, strSQLOrderInfo + Convert.ToString(order) + ";");
        }

        public DataTable GetShopInfoDataTable(int shopNumber)
        {
            return GetDataFromAccessDB(connectionStr, strSQLShop + Convert.ToString(shopNumber) + ";");
        }
    }
}
