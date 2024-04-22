using STOCK.Business;
using STOCK.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using System.Web.Services.Description;

namespace STOCK.Controllers
{
    public class StockController : Controller
    {
        string msg;
        List<object> excelData = new List<object>();
        Dictionary<DateTime, double> compositeIndex = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> PA_Bank = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> MT_Group = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> ZX_Group = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> HX_Group = new Dictionary<DateTime, double>();
        Dictionary<DateTime, double> TD_Group = new Dictionary<DateTime, double>();
        
        List<string> PA_labels_date = new List<string>();
        List<string> MT_labels_date = new List<string>();
        List<string> HX_labels_date = new List<string>();
        List<string> ZX_labels_date = new List<string>();
        List<string> TD_labels_date = new List<string>();

        List<double> PA_relativeReturn = new List<double>();
        List<double> MT_relativeReturn = new List<double>();
        List<double> HX_relativeReturn = new List<double>();
        List<double> ZX_relativeReturn = new List<double>();
        List<double> TD_relativeReturn = new List<double>();

        public StockController()
        {
            Unity.ReadExcel(excelData, compositeIndex, PA_Bank, MT_Group, HX_Group, ZX_Group, TD_Group);
        }
        // GET: Stock
        public ActionResult StockDisplay()
        {
            return View(new Data());
        }

        [HttpPost]
        public ActionResult StockDisplay(string[] stock, string startDate, string endDate)
        {
            if (stock == null)
            {
                msg = "请至少选择一只股票";
            }
            else
            {
                if (stock.Contains("PA"))
                {
                    Unity.GetDateLablesAndProfit(PA_labels_date, compositeIndex, PA_Bank, PA_relativeReturn, startDate, endDate, ref msg);
                }
                if (stock.Contains("MT"))
                {
                    Unity.GetDateLablesAndProfit(MT_labels_date, compositeIndex, MT_Group, MT_relativeReturn, startDate, endDate, ref msg);
                }
                if (stock.Contains("ZX"))
                {
                    Unity.GetDateLablesAndProfit(ZX_labels_date, compositeIndex, ZX_Group, ZX_relativeReturn, startDate, endDate, ref msg);
                }
                if (stock.Contains("HX"))
                {
                    Unity.GetDateLablesAndProfit(HX_labels_date, compositeIndex, HX_Group, HX_relativeReturn, startDate, endDate, ref msg);
                }
                if (stock.Contains("TD"))
                {
                    Unity.GetDateLablesAndProfit(TD_labels_date, compositeIndex, TD_Group, TD_relativeReturn, startDate, endDate, ref msg);
                }

                string PA_Data = Unity.GetDatas(PA_labels_date.ToArray(), PA_relativeReturn.ToArray());

                string MT_Data = Unity.GetDatas(MT_labels_date.ToArray(), MT_relativeReturn.ToArray());

                string HX_Data = Unity.GetDatas(HX_labels_date.ToArray(), HX_relativeReturn.ToArray());

                string ZX_Data = Unity.GetDatas(ZX_labels_date.ToArray(), ZX_relativeReturn.ToArray());

                string TD_Data = Unity.GetDatas(TD_labels_date.ToArray(), TD_relativeReturn.ToArray());

                return View(new Data { Message = msg, PA_Datas = PA_Data, MT_Datas = MT_Data, HX_Datas = HX_Data, ZX_Datas = ZX_Data, TD_Datas = TD_Data });
            }
            return View(new Data { Message = msg });
        }
    }
}