using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace AddonWebServiceXml
{
    public class ServiceLayer
    {
        public static dynamic Logout(string serviceLayerAddress, string session)
        {
            var httpWebRequest =
                (HttpWebRequest)WebRequest.Create(serviceLayerAddress +
                                                   string.Format(@"{0}", "/Logout"));//string logout, tira o cookie
                                                                                     //httpWebRequest.Headers.Add("B1S-ReplaceCollectionsOnPatch","true");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.AllowAutoRedirect = false;
            httpWebRequest.Timeout = 30 * 20000;
            httpWebRequest.ServicePoint.Expect100Continue = false;
            httpWebRequest.CookieContainer = new CookieContainer();
            ServicePointManager.ServerCertificateValidationCallback += BypassSslCallback;
            httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
                        new Cookie("B1SESSION", session));
            httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
                       new Cookie("ROUTEID", ".node1"));

            HttpWebResponse response1 = null;

            try
            {
                response1 = httpWebRequest.GetResponse() as HttpWebResponse;
            }
            catch (WebException ex)
            {
                using (WebResponse response = ex.Response)
                {
                    var httpResponse = (HttpWebResponse)response;

                    using (Stream data = response.GetResponseStream())
                    {
                        StreamReader sr = new StreamReader(data);

                        dynamic Dadoserror = Newtonsoft.Json.JsonConvert.DeserializeObject(sr.ReadToEnd());
                        var messageErro = Dadoserror.error.message;
                        //return Convert.ToString(messageErro["value"]);
                        throw new Exception("erro: " + Convert.ToString(messageErro["value"]));
                    }
                }

            }




            string responseContent = null;
            dynamic json = null;
            using (var reader = new StreamReader(response1.GetResponseStream(), Encoding.UTF8))
            {
                responseContent = reader.ReadToEnd();
                json = Newtonsoft.Json.JsonConvert.DeserializeObject(responseContent);

            }

            return "OK";
        }

        //string serviceLayerAddress, string User, string Pass, string DataBase
        public static dynamic Login()
        {
            var httpWebRequest =
                (HttpWebRequest)WebRequest.Create("https://hanab1:50000/b1s/v1" +
                                                    "/Login");//string logout, tira o cookie
                                                                                     //httpWebRequest.Headers.Add("B1S-ReplaceCollectionsOnPatch","true");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.AllowAutoRedirect = false;
            httpWebRequest.Timeout = 30 * 20000;
            httpWebRequest.ServicePoint.Expect100Continue = false;
            httpWebRequest.CookieContainer = new CookieContainer();

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                string body = "{\"CompanyDB\":\"SBO_COPROSUL_PRD\",\"Password\":\"B1admin@\",\"UserName\":\"manager\"}";

                streamWriter.Write(body);
            }


            ServicePointManager.ServerCertificateValidationCallback += BypassSslCallback;
            //httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
            //            new Cookie("B1SESSION", session));
            //httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
            //           new Cookie("ROUTEID", ".node1"));

            HttpWebResponse response1 = null;

            try
            {
                response1 = httpWebRequest.GetResponse() as HttpWebResponse;
            }
            catch (WebException ex)
            {
                using (WebResponse response = ex.Response)
                {
                    var httpResponse = (HttpWebResponse)response;

                    using (Stream data = response.GetResponseStream())
                    {
                        StreamReader sr = new StreamReader(data);

                        dynamic Dadoserror = Newtonsoft.Json.JsonConvert.DeserializeObject(sr.ReadToEnd());
                        var messageErro = Dadoserror.error.message;
                        //return Convert.ToString(messageErro["value"]);
                        throw new Exception("erro: " + Convert.ToString(messageErro["value"]));
                    }
                }

            }




            string responseContent = null;
            dynamic json = null;
            using (var reader = new StreamReader(response1.GetResponseStream(), Encoding.UTF8))
            {
                responseContent = reader.ReadToEnd();
                return json = Newtonsoft.Json.JsonConvert.DeserializeObject(responseContent);

            }

            //return "OK";
        }


        public static dynamic http(string url,  string type, string session = null, string body = null)
        {
            var httpWebRequest =
                (HttpWebRequest)WebRequest.Create(url);//string logout, tira o cookie
                                                              //httpWebRequest.Headers.Add("B1S-ReplaceCollectionsOnPatch","true");
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = type;
            httpWebRequest.AllowAutoRedirect = false;
            httpWebRequest.Timeout = 30 * 20000;
            httpWebRequest.ServicePoint.Expect100Continue = false;
            httpWebRequest.CookieContainer = new CookieContainer();

            if(body != null)
            {
                using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    streamWriter.Write(body);
                }
            }

            ServicePointManager.ServerCertificateValidationCallback += BypassSslCallback;

            if (session != null)
            {
                httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
                            new Cookie("B1SESSION", session));
                httpWebRequest.CookieContainer.Add(httpWebRequest.RequestUri,
                           new Cookie("ROUTEID", ".node1"));
            }


            HttpWebResponse response1 = null;

            try
            {
                response1 = httpWebRequest.GetResponse() as HttpWebResponse;
            }
            catch (WebException ex)
            {
                using (WebResponse response = ex.Response)
                {
                    var httpResponse = (HttpWebResponse)response;

                    using (Stream data = response.GetResponseStream())
                    {
                        StreamReader sr = new StreamReader(data);

                        dynamic Dadoserror = Newtonsoft.Json.JsonConvert.DeserializeObject(sr.ReadToEnd());
                        var messageErro = Dadoserror.error.message;
                        //return Convert.ToString(messageErro["value"]);
                        throw new Exception("erro: " + Convert.ToString(messageErro["value"]));
                    }
                }

            }




            string responseContent = null;
            dynamic json = null;
            using (var reader = new StreamReader(response1.GetResponseStream(), Encoding.UTF8))
            {
                responseContent = reader.ReadToEnd();
                return json = Newtonsoft.Json.JsonConvert.DeserializeObject(responseContent);

            }

            //return "OK";
        }


        private static bool BypassSslCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }



        //private static bool BypassSslCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        //{
        //    return true;
        //}

        //public static dynamic Login(string SL, string User, string Pass, string DataBase)
        //{
        //    RestResponse response1 = null;
        //    var client = new RestClient(SL + "/Login");

        //    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
        //    var request = new RestRequest(SL, Method.Post);

        //    request.AddHeader("Content-Type", "application/json");
        //    request.AddParameter("application/json", "{\n    \"CompanyDB\": \"" + DataBase + "\",\n    \"UserName\": \"" + User + "\",\n    \"Password\": \"" + Pass + "\"\n}", ParameterType.RequestBody);


        //    try
        //    {
                
        //        response1 = client.Execute(request);

        //    }
        //    catch (WebException ex)
        //    {
        //        using (WebResponse response = ex.Response)
        //        {
        //            var httpResponse = (HttpWebResponse)response;

        //            using (Stream data = response.GetResponseStream())
        //            {
        //                StreamReader sr = new StreamReader(data);

        //                dynamic Dadoserror = Newtonsoft.Json.JsonConvert.DeserializeObject(sr.ReadToEnd());
        //                var messageErro = Dadoserror.error.message;
        //                return Convert.ToString(messageErro["value"]);
        //            }
        //        }

        //    }

        //    if (response1.Content.Contains("erro"))
        //    {
        //        JObject jObject = JObject.Parse(response1.Content.ToString());
        //        string erro = jObject.SelectToken("error").ToString().Replace("\r", "").Replace("\n", "").Replace("\"", "").Replace("{", "").Replace("}", "");
        //        throw new Exception(erro);
        //        //throw new Exception("Erro" + response1.Content);
        //    }

        //    dynamic retorno = response1.Content;
        //    retorno = Newtonsoft.Json.JsonConvert.DeserializeObject(response1.Content);

        //    return retorno;



        //}
    }
}
