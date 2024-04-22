using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace STOCK.Models
{
    public class Data
    {
        public string Message { set; get; }
        public string PA_Datas { set; get; }
        public string MT_Datas { set; get; }
        public string HX_Datas { set; get; }
        public string ZX_Datas { set; get; }
        public string TD_Datas { set; get; }
        public Data()
        {
            PA_Datas = "[]";
            MT_Datas = "[]";
            HX_Datas = "[]";
            ZX_Datas = "[]";
            TD_Datas = "[]";
        }
    }

    public class Coord 
    {
        public string x;
        public double y;
    }
}