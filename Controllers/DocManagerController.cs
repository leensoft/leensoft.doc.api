using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using leensoft.doc.api.Library;

namespace leensoft.doc.api.Controllers
{
    public class DocManagerController : ApiController
    {
        public string Post([FromBody]PostModel data)
        {

            string result = string.Empty;
            //返回zip文件
            if (data.code == "zip")
            {
                JObject jsonObj = JObject.Parse(data.list);
                JToken record = jsonObj;
                DocModel model = null;
                foreach (JProperty jp1 in record)
                {
                    JArray jar = JArray.Parse(jp1.Value.ToString());
                    for (int i = 0; i < jar.Count; i++)
                    {
                        model = new DocModel();
                        model.template = data.template;
                        JToken record2 = jar[i];
                        foreach (JProperty jp2 in record2)
                        {
                            switch (jp2.Name)
                            {
                                case "name":
                                    model.name = jp2.Value.ToString();
                                    break;
                                case "doc":
                                    model.doc = jp2.Value.ToString();
                                    break;
                                default:
                                    break;
                            }
                        }
                        CreateDoc(model);
                    }
                }
                string rootUrl = System.Configuration.ConfigurationManager.AppSettings["WEBURL"];
                string path = System.Configuration.ConfigurationManager.AppSettings["TemplatePath"];
                ZipHelper.CreateZipFile(path + "build/", path + "build.zip");
                Directory.Delete(path + "build",true);
                //File.Delete(path + "build");
                //File.Delete(path + "build.zip");
                result = rootUrl + "\build\build.zip";
            }
            //返回doc文档
            if (data.code == "doc")
            {
                DocModel model = new DocModel();
                model.doc = data.list;
                model.template = data.template;
                model.name = data.name;
                result = CreateDoc(model);
            }
            return result;
            
        }


        //生成DOC文档
        private string CreateDoc(DocModel data)
        {
            Dictionary<string, string> dics1 = new Dictionary<string, string>();
            Dictionary<string, JArray> dics2 = new Dictionary<string, JArray>();
            try
            {
                string jsonData = data.doc;
                JObject jsonObj = JObject.Parse(jsonData);
                JToken record = jsonObj;
                CreateDocData(record, dics1, dics2);
            }
            catch (Exception e)
            {
                //FileTxtLogs.WriteLog(e.Message);
            }
            return DocHelper.CreateInvoice(data.template, data.name, dics1, dics2);   
        }

        //生成DOC数据
        private void CreateDocData(JToken record, Dictionary<string, string> dics1, Dictionary<string, JArray> dics2)
        {
            foreach (JProperty jp in record)
            {
                if (jp.Value.GetType().ToString() == "Newtonsoft.Json.Linq.JValue")
                {
                    if (!dics1.ContainsKey(jp.Name))
                    {
                        dics1.Add(jp.Name, jp.Value.ToString());
                    }

                }

                if (jp.Value.GetType().ToString() == "Newtonsoft.Json.Linq.JObject")
                {
                    JToken record2 = jp.Value;
                    CreateDocData(record2, dics1, dics2);
                }

                if (jp.Value.GetType().ToString() == "Newtonsoft.Json.Linq.JArray")
                {
                    JArray jar = JArray.Parse(jp.Value.ToString());
                    if (!dics2.ContainsKey(jp.Name))
                    {
                        dics2.Add(jp.Name, jar);
                    }
                }
            }
        }


    }

    public class PostModel
    {
        public string template { get; set; }
        public string name { get; set; }
        public string list { get; set; }
        public string code { get; set; }
        
    }

    public class DocModel
    {
        public string template { get; set; }
        public string name { get; set; }
        public string doc { get; set; }
    }
}
