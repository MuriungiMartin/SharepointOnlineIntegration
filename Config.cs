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
            public string message { get; set; }
            public object data { get; set; }





            public DefaultResponse(int status, string message, object data = null)
            {
                this.status = status;
                this.message = message;
                this.data = data;                
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
