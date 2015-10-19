using CSDNExtend.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace CSDNExtend.Common
{
    public class SqlHelper
    {
        private string connectionString;
        private string _cmd;

        public SqlHelper(string startDate,string endDate)
        {
            connectionString = "Server=10.168.172.243;Database=CSDNDB;User Id=sa;Password = Password01!; ";
            _cmd = string.Format("" +
                    "select " +
                    "[TeamName]," +
                    "case [CSSActionName] when 'Solution Delivered' then  'Yes' " +
                    "        when 'Answered' then 'Yes' else 'No'" +
                    "        end as [IsAnswered]" +
                    ",[Alias]" +
                    ",[ThreadLink]" +
                    ",[ThreadTitle]   ,[TechCategoryName]" +
                    ",[IssueTypeName]" +
                    ",[IRT]" +
                    ",[ThreadCreateTime]" +
                    ",[FirstReplyTime]" +
                    ",[Labors]" +
                    ",[ReplyNum]" +
                    ",[CSSActionName], " +
                    "case [IsReplied]" +
                    "when '1' then  'Yes' " +
                    "else 'No'" +
                    "end as [IsReplied]" +
                    ",[Diffcult]" +
                    ",[CustomerLookingFor]       " +
                    " FROM [CSDNDB].[dbo].[ThreadsDetail] {0}", " where ThreadCreateTime between '"+startDate+"' and '"+endDate+"'");
            ;
        }
        public List<ThreadDetail> GetCSDNThreads()
        {
            List<ThreadDetail> list = new List<ThreadDetail>();
           
            SqlConnection sql = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = _cmd;
            cmd.Connection = sql;
            cmd.CommandType = CommandType.Text;
            try
            {
                sql.Open();
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        // List<ThreadDetail> threads=reader.autoMa
                        while (reader.Read())
                        {
                            ThreadDetail thread = new ThreadDetail();
                            thread.Team = (reader["TeamName"] is DBNull) ? "" : reader["TeamName"].ToString();
                            thread.IsAnswered = (reader["IsAnswered"] is DBNull) ? "" : reader["IsAnswered"].ToString();
                            thread.Owner = (reader["Alias"] is DBNull) ? "" : (reader["Alias"]).ToString();
                            thread.Title = (reader["ThreadTitle"] is DBNull) ? "" : (reader["ThreadTitle"]).ToString();
                            thread.URL = (reader["ThreadLink"] is DBNull) ? "" : (reader["ThreadLink"]).ToString();
                            thread.TechCategory = (reader["TechCategoryName"] is DBNull) ? "" : (reader["TechCategoryName"]).ToString(); //reader["TechCategoryName"]
                            thread.IssueType = (reader["IssueTypeName"] is DBNull) ? "" : (reader["IssueTypeName"]).ToString();// reader["IssueTypeName"]
                            thread.IR = (reader["IRT"] is DBNull) ? "" : (reader["IRT"].ToString());
                            thread.CreateOn = (reader["ThreadCreateTime"] is DBNull) ? null : (DateTime?)reader["ThreadCreateTime"]; ;// (DateTime?)reader["ThreadCreateTime"];
                            thread.FirstReply = (reader["FirstReplyTime"] is DBNull) ? null : (DateTime?)reader["FirstReplyTime"];
                            thread.Labor = (reader["Labors"] is DBNull) ? null : (double?)(reader["Labors"]); //reader["UT"];
                            thread.Replies = (reader["ReplyNum"] is DBNull) ? 0 : Convert.ToInt16(reader["ReplyNum"] ?? 0);
                            thread.CssAction = (reader["CSSActionName"] is DBNull) ? "" : (reader["CSSActionName"]).ToString(); //reader["CSSActionName"]
                            thread.Replied = (reader["IsReplied"] is DBNull) ? "" : reader["IsReplied"].ToString();
                            thread.Difficulty = (reader["Diffcult"] is DBNull) ? "" : (reader["Diffcult"]).ToString(); //reader["Diffcult"]
                            thread.CustomLooking = (reader["Labors"] is DBNull) ? "" : (reader["CustomerLookingFor"] ?? "").ToString();// reader["CustomerLookingFor"];
                            thread.DayToAnswer = "";
                            thread.Contribution = "";
                            list.Add(thread);
                        }
                    }
                }
            }
            catch(Exception e)
            {
            }
            finally { sql.Close(); }

            return list;
        }
        public void GetTableValue(SqlDataReader reader,string property)
        {
            if (reader.IsDBNull(reader.GetOrdinal(property)))
            {

            }
        }
    }
}