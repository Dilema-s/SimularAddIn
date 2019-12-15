using System;
using System.Web;
using System.Web.Script.Serialization;
using SimularAddinWeb.QuantLib;

namespace SimularAddinWeb
{
    /// <summary>
    /// Handler for QuantLib
    /// </summary>
    public class Handler : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentEncoding = System.Text.Encoding.UTF8;
            context.Response.ContentType = "text/html";

            var strResult = string.Empty;

            switch (context.Request.QueryString["RequestType"].Trim())
            {
                case "Derivative":
                    strResult = this.Derivative(
                        Convert.ToDouble(context.Request["number1"]));
                    break;
            }
            context.Response.Write(strResult);
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

        public string Derivative(double number1)
        {
            double total = 0;

            TestClass sampleLib= new TestClass();
            total = sampleLib.Derivative(number1);

            var jsonData = new
            {
                Total = total
            };
            JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();

            return javaScriptSerializer.Serialize(jsonData);
        }
    }
}