using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace leensoft.doc.api.Library
{
    public partial class BaseModel
    {
        //protected static OperContext oc = OperContext.CurrentContext;
    }

    public class Reslut
    {
        public string code { get; set; }
        public string msg { get; set; }
        public virtual DataModel data { get; set; }
    }


    public class DataModel
    {
               
    }


    

}