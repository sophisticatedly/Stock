using Microsoft.Ajax.Utilities;
using Newtonsoft.Json;
using OfficeOpenXml;
using STOCK.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Windows.Interop;

namespace STOCK.Business
{
    public class Unity
    {
        public static void ReadExcel(List<object> list, Dictionary<DateTime, double> compositeIndex, Dictionary<DateTime, double> PA_Bank, Dictionary<DateTime, double> MT_Group, Dictionary<DateTime, double> HX_Group, Dictionary<DateTime, double> ZX_Group, Dictionary<DateTime, double> TD_Group)
        {
            // 读取Excel文件
            string excelFile = AppDomain.CurrentDomain.BaseDirectory + @"App_Data\work.xlsx";
            FileInfo fileInfo = new FileInfo(excelFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // 获取第2个工作表
                object cellValue;                                                        // 遍历行和列，将数据存储到List中
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {                        
                        if (col == 1 && row != 1)
                        {
                            cellValue = worksheet.Cells[row, col].Value; // 读取单元格的值
                            DateTime date = DateTime.FromOADate((double)cellValue);
                            list.Add(date.ToString("yyyy-MM-dd"));
                            continue;
                        }
                        cellValue = worksheet.Cells[row, col].Value?.ToString(); // 读取单元格的值
                        list.Add(cellValue); // 将数据添加到List中
                    }
                }
                DateTime dateTime = DateTime.Now;
                double compositeIndexOneday = 0;
                double PA_BankIndex = 0;
                double MT_GroupIndex = 0;
                double HX_GroupIndex = 0;
                double ZX_GroupIndex = 0;
                double TD_GroupIndex = 0;
                //添加数据到compositeIndex
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (col == 1 && row != 1 && worksheet.Cells[row, col].Value != null )
                        {
                            cellValue = worksheet.Cells[row, col].Value; // 读取单元格的值
                            dateTime = DateTime.FromOADate((double)cellValue);
                        }
                        if(col == 2 && row != 1 && worksheet.Cells[row, col].Value != null )
                        {
                            compositeIndexOneday = (double)worksheet.Cells[row, col].Value;
                        }
                        if(row != 1 && col == 2 && worksheet.Cells[row, col].Value != null )
                        {
                            compositeIndex.Add(dateTime, compositeIndexOneday);
                        }
                        if (row != 1 && col == 3 && worksheet.Cells[row, col].Value != null)
                        {
                            PA_BankIndex = (double)worksheet.Cells[row, col].Value;
                            PA_Bank.Add(dateTime, PA_BankIndex);
                        }
                        if(row != 1 && col == 4 && worksheet.Cells[row, col].Value != null)
                        {
                            MT_GroupIndex = (double)worksheet.Cells[row, col].Value;
                            MT_Group.Add(dateTime, MT_GroupIndex);
                        }
                        if (row != 1 && col == 5 && worksheet.Cells[row, col].Value != null)
                        {
                            ZX_GroupIndex = (double)worksheet.Cells[row, col].Value;
                            ZX_Group.Add(dateTime, ZX_GroupIndex);
                        }
                        if (row != 1 && col == 6 && worksheet.Cells[row, col].Value != null)
                        {                            
                            HX_GroupIndex = (double)worksheet.Cells[row, col].Value;
                            HX_Group.Add(dateTime, HX_GroupIndex);                                                        
                        }
                        if (row != 1 && col == 7 && worksheet.Cells[row, col].Value != null)
                        {                            
                            TD_GroupIndex = (double)worksheet.Cells[row, col].Value;
                            TD_Group.Add(dateTime, TD_GroupIndex);                            
                        }
                    }
                }
            }
        }

        public static void GetDateLablesAndProfit(List<string> labels_date, Dictionary<DateTime, double> compositeIndex, Dictionary<DateTime, double> stocksOfBank, List<double> relativeReturn, string startDate, string endDate, ref string msg)
        {
            try
            {
                if (endDate.IsNullOrWhiteSpace() || startDate.IsNullOrWhiteSpace())
                {
                    msg = "开始日期或结束日期没有输入";
                }
                else if (DateTime.Parse(endDate) < DateTime.Parse(startDate))
                {
                    msg = "结束时间不能早于开始时间";
                }
                else
                {
                    int days = (int)DateTime.Parse(endDate).Subtract(DateTime.Parse(startDate)).TotalDays + 1;

                    for (int i = 0; i < days; i++)
                    {
                        if (compositeIndex.ContainsKey(DateTime.Parse(startDate).AddDays(i)) && stocksOfBank.ContainsKey(DateTime.Parse(startDate).AddDays(i)))
                        {
                            labels_date.Add(DateTime.Parse(startDate).AddDays(i).ToString("yyyy-MM-dd"));
                        }
                    }
                    string[] labels_datetime = labels_date.ToArray();
                    if (CheckDate(stocksOfBank, startDate, endDate))
                    {
                        relativeReturn.Add(1);
                    }
                    for (int i = 1; i < labels_datetime.Length; i++)
                    {
                        if (labels_datetime.Length > 0)
                        {
                            double compositeRateChanged = (compositeIndex[DateTime.Parse(labels_datetime[i])] - compositeIndex[DateTime.Parse(labels_datetime[i - 1])]) / compositeIndex[DateTime.Parse(labels_datetime[i - 1])];
                            double stockRateChanged = (stocksOfBank[DateTime.Parse(labels_datetime[i])] - stocksOfBank[DateTime.Parse(labels_datetime[i - 1])]) / stocksOfBank[DateTime.Parse(labels_datetime[i - 1])];
                            double relativeReturnOneDay = stockRateChanged - compositeRateChanged + relativeReturn[i - 1];
                            relativeReturn.Add(Math.Round(relativeReturnOneDay, 2));
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                msg = exception.Message;
            }           
        }

        private static bool CheckDate(Dictionary<DateTime, double> stocksOfBank, string startDate, string endDate)
        {
            bool overlap = false;
            if(DateTime.Parse(startDate) <= stocksOfBank.LastOrDefault().Key)
            {
                overlap = true;
            }
            return overlap;
        }

        public static string GetDatas(string[] dates, double[] vales)
        {
            int length = dates.Length;
            List<Coord> coordList = new List<Coord>();
            for (int i = 0; i < length; i++)
            {
                coordList.Add(new Coord { x = dates[i], y = vales[i] });
            }
            return JsonConvert.SerializeObject(coordList);
        }
    }
}