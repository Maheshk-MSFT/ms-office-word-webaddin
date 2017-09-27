using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

namespace PeriodicTableAppsWeb
{
    /// <summary>
    /// Temp page for interaction
    /// </summary>
    public partial class temppage : System.Web.UI.Page
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Page_Load(object sender, EventArgs e)
        {
            string inputtype = Convert.ToString(Request.QueryString["input"]);

            if (inputtype == null)
                return;

            var _url = "http://www.webservicex.net/periodictable.asmx";
            var _action = "http://www.webserviceX.NET/GetAtomicNumber";
            XmlDocument soapEnvelopeXml = CreateSoapEnvelope(inputtype);
            HttpWebRequest webRequest = CreateWebRequest(_url, _action);
            InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);
            IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
            asyncResult.AsyncWaitHandle.WaitOne();
            string soapResult;
            string returnvalue = "";

            using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
            {
                using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                {
                    soapResult = rd.ReadToEnd();
                }

                XmlDocument resultXml = new XmlDocument();
                XmlNamespaceManager mgr = new XmlNamespaceManager(resultXml.NameTable);
                mgr.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
                mgr.AddNamespace("multi", "http://www.webserviceX.NET");
                resultXml.LoadXml(soapResult);
                XmlNodeList nodeList = resultXml.SelectNodes("//multi:GetAtomicNumberResult", mgr);

                string result = string.Empty;

                if (nodeList.Count > 0)
                {
                    XmlNode nd = nodeList[0];

                    string resultstring = nd.InnerText;

                    resultstring = resultstring.Replace("<NewDataSet>", "").Replace("<Table>", "").Replace("</Table>", "").Replace("</NewDataSet>", "");
                    resultstring = resultstring.Replace("<NewDataSet>", "").Replace("<Table>", "").Replace("</Table>", "").Replace("</NewDataSet>", "");
                    resultstring = resultstring.Replace("</AtomicNumber>", "").Replace("</ElementName>", "").Replace("</Symbol>", "");
                    resultstring = resultstring.Replace("</AtomicWeight>", "").Replace("</BoilingPoint>", "").Replace("</IonisationPotential>", "");
                    resultstring = resultstring.Replace("</EletroNegativity>", "").Replace("</AtomicRadius>", "").Replace("</MeltingPoint>", "").Replace("</Density>", "");

                    var split = resultstring.Split(new[] { "<" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string str in split)
                    {
                        if (str.Contains(">"))
                        {
                            var keyval = str.Trim().Split(new[] { ">" }, StringSplitOptions.RemoveEmptyEntries);
                            returnvalue += "~" + keyval[0].ToString() + "," + keyval[1].ToString();
                        }
                    }
                }
                Response.Write(returnvalue);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static XmlDocument CreateSoapEnvelope(string elementname)
        {
            XmlDocument soapEnvelop = new XmlDocument();
            soapEnvelop.LoadXml(@"<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""><SOAP-ENV:Body><GetAtomicNumber  xmlns=""http://www.webserviceX.NET""> <ElementName>" + elementname + "</ElementName></GetAtomicNumber></SOAP-ENV:Body></SOAP-ENV:Envelope>");
            return soapEnvelop;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="url"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        private static HttpWebRequest CreateWebRequest(string url, string action)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("SOAPAction", action);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="soapEnvelopeXml"></param>
        /// <param name="webRequest"></param>
        private static void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest)
        {
            using (Stream stream = webRequest.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }
        }
    }
}