using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SharepointOnlineIntegration
{
    public class Config
    {
        public struct DefaultResponse
        {
            public int status { get; set; }
            public int resNumber { get; set; }
            public bool res { get; set; }
            public string message { get; set; }
            public string return_value { get; set; }
            public string retvalue { get; set; }
            public string dashvalue { get; set; }

            public object data { get; set; }





            public DefaultResponse(int status, int resNumber, string message, object data = null, string return_value = null, bool res = false, string dashvalue = null)
            {
                this.status = status;
                this.resNumber = resNumber;
                this.message = message;
                this.data = data;
                this.return_value = return_value;
                this.retvalue = return_value;
                this.dashvalue = dashvalue;
                this.res = res;




            }
            public struct JSONArrayResponse
            {
                public int status { get; set; }
                public JArray response { get; set; }



                public JSONArrayResponse(int status, JArray message)
                {
                    this.status = status;
                    this.response = message;
                }
            }
        }
    }
}
