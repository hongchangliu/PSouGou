using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.IO;
using System.Collections;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Data;

namespace PSouGou
{
    class Program
    {
        static void Main(string[] args)
        {
            DirectoryInfo folder = new DirectoryInfo("output/Batchs/SChengDu_001");
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("序号", typeof(int));
            dt.Columns.Add("城市", typeof(String));
            dt.Columns.Add("主点类型", typeof(String));
            dt.Columns.Add("补充主点名称", typeof(String));
            dt.Columns.Add("现有主点名称", typeof(String));
            dt.Columns.Add("现有主点搜狗ID", typeof(String));
            dt.Columns.Add("别名", typeof(String));
            dt.Columns.Add("现有主点是否结构化正确", typeof(String));
            dt.Columns.Add("地址", typeof(String));
            dt.Columns.Add("大类", typeof(String));
            dt.Columns.Add("小类", typeof(String));
            dt.Columns.Add("电话", typeof(String));
            dt.Columns.Add("X", typeof(String));
            dt.Columns.Add("Y", typeof(String));
            dt.Columns.Add("备注", typeof(String));
            dt.Columns.Add("标注员", typeof(String));
            foreach (FileInfo file in folder.GetFiles("*.xml"))
            {
                DataRow dr = dt.NewRow();
                Console.WriteLine("开始读取：" + file.Name);
                XElement xdoc = XElement.Load(file.FullName);

                var ad = from column in xdoc.Descendants("Row").Elements("Column")
                         select new
                         {
                             name = column.Attribute("name").Value,
                             val = column.Element("Datas").Elements().Last().Value
                         };

                //标注员
                String optr = xdoc.Descendants("Stats").Elements().First().Element("Operator").Value.Trim();


                foreach (var a in ad)
                {
                    if (a == null)
                    {
                        Console.WriteLine("存在非法数据");
                        continue;
                    }


                    if ("序号".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }
                    if ("城市".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    if ("主点类型".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim().Substring(1);
                    }
                    if ("补充主点名称".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }
                    if ("现有主点名称".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }
                    if ("现有主点搜狗ID".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    if ("别名".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }


                    if ("现有结构化是否正确".Equals(a.name))
                    {
                        dr["现有主点是否结构化正确"] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    if ("地址".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }


                    if ("大类".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    if ("小类".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    if ("电话".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }
                    if (dr["主点类型"] != null)
                    {
                        bool b = dr["主点类型"].ToString().Contains("缺失主点") || dr["主点类型"].ToString().Contains("缺失主点关系错误");
                        if (b)
                        {
                            dr["X"] = "X";
                            dr["Y"] = "Y";
                        }

                    }

                    if ("备注".Equals(a.name))
                    {
                        dr[a.name] = a.val == null || "NULL".Equals(a.val) ? null : a.val.Trim();
                    }

                    dr["标注员"] = optr;

                }
                dt.Rows.Add(dr);
            }
            dt.DefaultView.Sort = "序号 ASC";

            dt = dt.DefaultView.ToTable();
            ExcelEdit excelHelper = new ExcelEdit();
            excelHelper.Open(Environment.CurrentDirectory + "\\output\\Batchs\\templete.xlsx");
            Worksheet ws = excelHelper.GetSheet("Sheet1");
            excelHelper.InsertTable(dt, ws, 2, 1);
            bool b1 = excelHelper.SaveAs(Environment.CurrentDirectory + "\\output\\Save\\PSouGou.xlsx");
            excelHelper.Close();
            Console.WriteLine("写入完成!!!!!按回车键退出");
            Console.Read();

        }




    }
}
