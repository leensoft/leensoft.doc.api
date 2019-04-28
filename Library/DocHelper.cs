using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Learn.Utility;
using System.IO;

using Novacode;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace leensoft.doc.api.Library
{
    public class DocHelper
    {
        public static string CreateInvoice(string templateDoc,string buildDoc,Dictionary<string, string> dics1, Dictionary<string, JArray> dics2)
        {
            Console.WriteLine("\tCreateInvoice()");
            DocX g_document;
            string buildPath = "";
            try
            {
                string path = System.Configuration.ConfigurationManager.AppSettings["TemplatePath"];
                g_document = CreateInvoiceFromTemplate(DocX.Load(path + templateDoc), dics1, dics2);

                string rootUrl = System.Configuration.ConfigurationManager.AppSettings["WEBURL"];
                buildPath = rootUrl +"build/"+buildDoc;
                if (!Directory.Exists(path + "build/"))
                {
                    Directory.CreateDirectory(path + "build/");
                }
                g_document.SaveAs(path + "build/" + buildDoc);
              

            }
            catch(Exception e)
            {
                FileTxtLogs.WriteLog(e.Message);
               
            }
            return buildPath;
        }

        // Create an invoice for a factitious company called "The Happy Builder".
        private static DocX CreateInvoiceFromTemplate(DocX template, Dictionary<string, string> dics1, Dictionary<string, JArray> dics2)
        {
            //填充属性值
            FillAttrVal(dics1, ref template);

            //填充模板表格   
            for (int i = 0; i < template.Tables.Count; i++)
            {
                FillTable(ref dics2, ref template, i);
            }
            return template;
        }


        private static void FillAttrVal(Dictionary<string, string> dics, ref DocX document)
        {
            foreach (KeyValuePair<string, string> kvp in dics)
            {
                document.AddCustomProperty(new CustomProperty(kvp.Key, kvp.Value));
            }

        }

        private static void FillTable(ref Dictionary<string, JArray> dics, ref DocX document, int tbnum)
        {
            int cell = 0;
            //foreach (KeyValuePair<string, JArray> kvp in dics)
            Table fillTable = document.Tables[tbnum];
            for (int index = 0; index < dics.Count; index++)
            {
                KeyValuePair<string, JArray> kvp = dics.ElementAt(index);
                if (fillTable.Rows.Count == 2)
                {
                    string pname = ((Novacode.DocXElement)(fillTable.Rows[1].Cells[0])).Xml.Value;

                    if (pname == kvp.Key)
                    {

                        if (kvp.Value.Count == 0)
                        {
                            Row row = fillTable.InsertRow();
                            row.MergeCells(0, row.Cells.Count - 1);                            
                            Paragraph cell_paragraph = row.Cells[0].Paragraphs[0];
                            cell_paragraph.InsertText(pname+"暂无信息", false);
                            row.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                            row.Height = fillTable.Rows[0].Height;
                        }
                        else
                        {
                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                Row row = fillTable.InsertRow();
                                cell = 0;
                                JToken record = kvp.Value[i];

                                foreach (JProperty jp in record)
                                {
                                    row.Cells[cell].Width = fillTable.Rows[0].Cells[cell].Width;
                                    Paragraph cell_paragraph = row.Cells[cell].Paragraphs[0];
                                    cell_paragraph.InsertText(jp.Value.ToString(), false);
                                    cell++;
                                }

                            }

                        }
                        fillTable.Rows[1].Remove();
                        dics.Remove(pname);

                    }

                }
                else if (fillTable.Rows.Count == 3)
                {
                    string pname = ((Novacode.DocXElement)(fillTable.Rows[2].Cells[0])).Xml.Value;

                    if (pname == kvp.Key)
                    {
                        if (kvp.Value.Count == 0)
                        {
                            fillTable.Design = TableDesign.None;
                            Row row = fillTable.InsertRow(fillTable.Rows[2]);
                            Paragraph cell_paragraph = row.Cells[0].Paragraphs[0];
                            cell_paragraph.InsertText("暂无信息", false);
                            row.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                        }
                        else
                        {

                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                Row row1 = fillTable.InsertRow(fillTable.Rows[0]);
                                Row row2 = fillTable.InsertRow(fillTable.Rows[1]);
                                cell = 0;
                                JToken record = kvp.Value[i];
                                foreach (JProperty jp in record)
                                {
                                    int cellCount = row1.Cells.Count;
                                    if (cell < (cellCount / 2))
                                    {
                                        //string pname1 = ((Novacode.DocXElement)(fillTable.Rows[2 * (i + 1) - 1].Cells[2 * (cell + 1) - 1])).Xml.Value;
                                        Paragraph cell_paragraph = row1.Cells[2 * (cell + 1) - 1].Paragraphs[0];
                                        cell_paragraph.InsertText(jp.Value.ToString(), false);
                                    }
                                    else
                                    {
                                        cell = 0;
                                        Paragraph cell_paragraph = row2.Cells[2 * (cell + 1) - 1].Paragraphs[0];
                                        cell_paragraph.InsertText(jp.Value.ToString(), false);
                                    }
                                    cell++;
                                }
                            }

                        }
                        dics.Remove(pname);
                        fillTable.Rows[0].Remove();
                        fillTable.Rows[1].Remove();
                        fillTable.Rows[0].Remove();

                    }

                }


                fillTable.AutoFit = AutoFit.Contents;
                // Center the Table
                fillTable.Alignment = Alignment.center;

            }

        }
    }
}