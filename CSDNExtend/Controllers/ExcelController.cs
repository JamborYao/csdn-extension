using CSDNExtend.Common;
using CSDNExtend.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web.Http;

namespace CSDNExtend.Controllers
{
    public class ExcelController : ApiController
    {
        public void Get(string startDate,string endDate)
        {
            SqlHelper sql = new SqlHelper(startDate,endDate);
            List<ThreadDetail> list = sql.GetCSDNThreads();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            Worksheet ws = (Worksheet)wb.Worksheets[1];
            int r = 1;
            foreach (ThreadDetail thread in list)
            {
               
                //int c = 1;
                string url = "";
                Type t = thread.GetType();
                ws.Cells[r, 1] = thread.Team;
                ws.Cells[r, 2] = thread.IsAnswered;
                ws.Cells[r, 3] = thread.Owner;
                ws.Cells[r, 4] = thread.Title;
                ws.Hyperlinks.Add(ws.Cells[r, 4],thread.URL);
                ws.Cells[r, 5] = thread.TechCategory;
                ws.Cells[r, 6] = thread.IssueType;
                ws.Cells[r, 7] = thread.IR;
                ws.Cells[r, 8] = thread.CreateOn;
                ws.Cells[r, 8].NumberFormat = "yyyy/m/d h:mm"; 
                ws.Cells[r, 9] = thread.FirstReply;
                ws.Cells[r, 9].NumberFormat = "yyyy/m/d h:mm";
                ws.Cells[r, 10] = thread.Labor;
                ws.Cells[r, 11] = thread.Replies;
                ws.Cells[r, 12] = thread.CssAction;
                ws.Cells[r, 13] = thread.Replied;
                ws.Cells[r, 14] = thread.Difficulty;
                ws.Cells[r, 15] = thread.CustomLooking;
                ws.Cells[r, 16] = thread.DayToAnswer;
                ws.Cells[r, 17] = thread.Contribution;
                  
                //foreach (var property in t.GetProperties())
                //{
                //    var attribute = property.GetCustomAttributes(true);
                //    if (attribute.Count() > 0)
                //    {
                //        if ((string)attribute[0].GetType().Name == "IsURLAttribute")
                //        {
                //            url = (string)property.GetValue(thread, null);
                //            continue;
                //        }
                //        if ((string)attribute[0].GetType().Name == "IsTitleAttribute")
                //        {
                //            ws.Hyperlinks.Add(ws.Cells[r, c], url);
                //        }

                //    }
                //    ws.Cells[r, c] = property.GetValue(thread, null);
                //    FormatExcelCell(property, ws.Cells[r, c]);
                //    c++;
                //}
                r++;
            }
            wb.Worksheets[1].Name = "CSDN threads";//Renaming the Sheet1 to MySheet
            wb.SaveAs("d:\\CSDNReport-"+DateTime.Now.ToString("yyyy-MM-dd")+".xlsx");
            wb.Close();
            xlApp.Quit();
           
        }
        // POST api/values
        public void Post([FromBody]string value)
        {
            //SqlHelper sql=new SqlHelper();
            //List<ThreadDetail> list= sql.GetCSDNThreads();
            //Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            //Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            //Worksheet ws = (Worksheet)wb.Worksheets[1];
            //foreach(ThreadDetail thread in list)            
            //{
            //    int r = 1;
            //    int c = 1;
            //    string url = "";
            //    Type t = thread.GetType();
            //    foreach (var property in t.GetProperties())
            //    {
            //        var attribute = property.GetCustomAttributes(true);
            //        if (attribute.Count() > 0)
            //        {
            //            if ((string)attribute[0].GetType().Name == "IsURLAttribute")
            //            {
            //                url = (string)property.GetValue(thread, null);
            //                continue;
            //            }                     
            //            if ((string)attribute[0].GetType().Name == "IsTitleAttribute")
            //            {
            //                ws.Hyperlinks.Add(ws.Cells[r, c], url);
            //            }
                        
            //        }
            //        ws.Cells[r, c] = property.GetValue(thread, null);                   
            //        FormatExcelCell(property, ws.Cells[r,c]);
            //        c++;
            //    }
            //    r++;
            //}
            //wb.Worksheets[1].Name = "MySheet";//Renaming the Sheet1 to MySheet
            //wb.SaveAs("d:\\Testing.xlsx");
            //wb.Close();
            //xlApp.Quit();
        }
        public static void FormatExcelCell(PropertyInfo property, Range cell)
        {
            string typeName = "";
            var nullableType = Nullable.GetUnderlyingType(property.PropertyType);
            bool isNullableType = nullableType != null;
            if (isNullableType)
                typeName = nullableType.FullName;
            else
                typeName= property.PropertyType.FullName;

            switch (typeName)
            {
                case "System.Int":
                    cell.NumberFormat = "0";
                    break;
                case "System.DateTime":
                    cell.NumberFormat = "yyyy/m/d h:mm";
                    break;
                case "System.Double":
                    cell.NumberFormat = "0.00";
                    break;
            }
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
            thread.FirstReply = null;
            thread.Labor = 24;
            thread.Replies = 3;
            thread.CssAction = "Answered";
            thread.Replied = "Yes";
            thread.Difficulty = "test";
            thread.CustomLooking = "test";
            thread.DayToAnswer = "test";
            return thread;
        }

        public void DownloadExcel()
        {
         
        }
    }
}
