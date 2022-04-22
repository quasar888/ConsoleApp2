using System;
using System.Data;
using System.IO;
using System.Collections.Generic;
using static Axpo.PowerService;
using System.Threading;
using System.Configuration;

namespace ConsoleApp2
{
    internal class Program
    {
        private static DataSet dtSet;
        static void Main(string[] args)
        {
            Console.WriteLine("Would you like setting the path (Create the csv file in a folder) for csv outfile or let the file path by default : Type Y or N and tape enter...");
            var response = Console.ReadLine();
            var response2 = "";
            if (response == "Y" || response == "y")
            {
                Console.WriteLine("Write the path you want a set file the csv outfile : ");
                response2 = Console.ReadLine();
            }
            Console.WriteLine("Would you like to specified tne interval time when the app will running or let the default interval running +/- 1 minutes : " +
                "Type Y or N and tape enter...");
            var response3 = Console.ReadLine();
            var response4 = "";

            if (response3 == "Y" || response3 == "y")
            {
                Console.WriteLine(" Please speficied your interval time in an integer minutes and tape enter... ");
                response4 = Console.ReadLine();
            }

            while(true)
            { 
            DataTable dataTable = new DataTable();
            DataColumn dtColumn;
            DataRow myDataRow;

            dtColumn = new DataColumn();
            dtColumn.DataType = typeof(DateTime);
            dtColumn.ColumnName = "Period";
            dataTable.Columns.Add(dtColumn);

            dtColumn = new DataColumn();
            dtColumn.DataType = typeof(double);
            dtColumn.ColumnName = "Volume";
            dataTable.Columns.Add(dtColumn);
            dtSet = new DataSet();
            dtSet.Tables.Add(dataTable);

            Axpo.PowerService powerService = new Axpo.PowerService();
            
            var date = DateTime.Now.Day;
            
            var timeStart = new DateTime();

            if (DateTime.Now.Day > 1)
            {

                timeStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, date - 1, 23, 0, 0);
            }
            else
            {
                if (DateTime.Now.Month != 1)
                {
                    timeStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, date - 1, 23, 0, 0);
                }
                else
                {
                    timeStart = new DateTime(DateTime.Now.Year - 1, DateTime.Now.Month - 1, date - 1, 23, 0, 0);
                }
            }
            int countperiod = 0;
            var volumeTotal_Hour = new double[24];
            Console.WriteLine("Program loading please wait....");
            while (countperiod < 24)
            {
                int count = 0;
                try
                {
                    foreach (var item in powerService.GetTradesAsync(timeStart.AddHours(countperiod)).Result)
                    {                        
                        foreach (var item2 in item.Periods)
                        {                            
                            if (count < 23)
                            {
                                count++;
                            }
                            else
                            {
                                count = 0;
                            }
                            volumeTotal_Hour[countperiod] += Math.Ceiling(item.Periods[count].Volume);
                        }                        
                    }
                    countperiod++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Program loading please wait....");
                }
            }
            for (int i = 0; i < 24; i++)
            {
                myDataRow = dataTable.NewRow();
                myDataRow["Period"] = timeStart.AddHours(i);
                myDataRow["Volume"] = volumeTotal_Hour[i];
                dataTable.Rows.Add(myDataRow);
            }
            var path = (response == "Y" || response == "y") ? response2 : ConfigurationManager.AppSettings["Path"] + DateTime.Now.ToString().Replace("/", "_").Replace(":", "") + ".csv";
            SaveDataSetAsExcel(dataTable, path);

            if (response4 != "")
            {
                Thread.Sleep(60000 * Convert.ToInt16(response4));
            }
            else
            {
                Thread.Sleep(60000);
            }
            }
        }

        public static void SaveDataSetAsExcel(DataTable dataTable, string exceloutFilePath)
        {
            try
            {                
                StreamWriter sw = new StreamWriter(exceloutFilePath, false);                
                int iColCount = dataTable.Columns.Count;
                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write(dataTable.Columns[i]);
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
                foreach (DataRow dr in dataTable.Rows)
                {
                    for (int i = 0; i < iColCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write(dr[i].ToString());
                        }
                        if (i < iColCount - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
