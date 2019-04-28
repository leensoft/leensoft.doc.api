using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;

namespace leensoft.doc.api.Library
{
    public class XmlHelper
    {
        private string xmlPath = System.Web.HttpContext.Current.Server.MapPath("/App_Data/smscode.xml");

        public List<XmlEntity> GetXmlNode()
        {
            List<XmlEntity> list = new List<XmlEntity>();
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            XmlNodeList nodeList = doc.ChildNodes[1].ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                var xmlEntity = new XmlEntity();
                xmlEntity.XmlValue = node.Attributes["value"].Value;
                xmlEntity.XmlName = node.Attributes["name"].Value;
                list.Add(xmlEntity);
            }
            return list;
        }

        public bool IsNodeExist(string phone)
        {   
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);    
            XmlNodeList nodeList = doc.ChildNodes[1].ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes["name"].Value == phone)
                {
                    return true;
                }
            }
            return false;
        }

        public bool IsNodeExist(string phone,string code)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            XmlNodeList nodeList = doc.ChildNodes[1].ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes["name"].Value == phone && node.Attributes["value"].Value == code)
                {
                    return true;
                }
            }
            return false;
        }

        public void SetXmlNode(string name, string v)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);           

            XmlElement xmlElement = doc.CreateElement("item");
            //添加属性
            xmlElement.SetAttribute("name", name);
            xmlElement.SetAttribute("value", v);
            //将节点加入到指定的节点下
            XmlNode xml = doc.DocumentElement.PrependChild(xmlElement);
            doc.Save(xmlPath);
        }

        public void EditXmlNode(string name, string v)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            XmlNodeList nodeList = doc.ChildNodes[1].ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes["name"].Value == name)
                {
                    node.Attributes["value"].Value = v;
                }
            }
            doc.Save(xmlPath);       
           
        }
    }

    public class XmlEntity
    {
        public string XmlName {get;set;}
        public string XmlValue {get;set;}
    }
}
