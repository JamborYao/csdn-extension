using CSDNExtend.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web.Http;

namespace CSDNExtend.Controllers
{
    public class ExcelController : ApiController
    {
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }
        // POST api/values
        public void Post([FromBody]string value)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            Worksheet ws = (Worksheet)wb.Worksheets[1];

            for (int r = 1; r < 5; r++) //r stands for ExcelRow and c for ExcelColumn
            {
                int c = 1;
                ThreadDetail thread = GetThread();
                string url = "";
                Type t = thread.GetType();
                foreach (var property in t.GetProperties())
                {
                    var attribute = property.GetCustomAttributes(true);
                    if (attribute.Count() > 0)
                    {
                        if ((string)attribute[0].GetType().Name == "IsURLAttribute")
                        {
                            url = (string)property.GetValue(thread, null);
                            continue;
                        }                     
                        if ((string)attribute[0].GetType().Name == "IsTitleAttribute")
                        {
                            ws.Hyperlinks.Add(ws.Cells[r, c], url);
                        }
                        
                    }
                    if (property.GetType().ToString() == "System.DateTime")
                    {
                        //Range rg = (Range)ws.Cells[1, 1];
                        //rg.EntireColumn.NumberFormat = "MM/DD/YYYY";
                        ws.Cells[r, c].NumberFormat = "m/d/yy";
                    }
                    ws.Cells[r, c] = property.GetValue(thread, null);
                    c++;
                }
            }



            wb.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet

            wb.SaveAs("d:\\Testing.xlsx");

            wb.Close();

            xlApp.Quit();
        }
        public int GetPorpertyNumber()
        {
            return 13;
        }

        public ThreadDetail GetThread()
        {
            ThreadDetail thread = new ThreadDetail();
            thread.Team = "1";
            thread.IsAnswered = "Yes";
            thread.Owner = "v-jayao";
            thread.Title = "ADO.NET classes in .NET 2.0 / 3.5哪个在windows azure不被支持？为保证我的程序在windows azure兼容，对于这些版本需要做哪些处理方法";
            thread.URL = "http://ask.csdn.net/questions/168465";
            thread.TechCategory = "General Discussion";
            thread.IssueType = "Mooncake feature - General Discussion";
            thread.IR = "359";
            thread.CreateOn = Convert.ToDateTime("2015-03-04 17:08:00.000");
            thread.FirstReply = "null";
            thread.Labor = "24";
            thread.Replies = "3";
            thread.CssAction = "Answered";
            thread.Replied = "2";
            thread.Difficulty = "test";
            thread.CustomLooking = "test";
            thread.DayToAnswer = "test";
            return thread;
        }
    }
}
