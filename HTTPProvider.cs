using System;
using System.Collections;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web.Script.Serialization;

namespace Project.Utility {
    public class HTTPProvider {
        /// <summary>  
        /// 证书路径  
        /// </summary>  
        public static String certFilePath = ConfigurationManager.AppSettings["certFilePath"];
        /// <summary>  
        /// 证书口令  
        /// </summary>  
        public static String certFilePwd = ConfigurationManager.AppSettings["certFilePwd"];
        
        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="para"></param>
        public static string HttpGet(string url, string para, string contentType) {
            try {
                url = url + "?" + para;
                HttpWebRequest request = (System.Net.HttpWebRequest)WebRequest.Create(url);
                request.ContentType = contentType;
                request.Method = "get";
                request.Timeout = 6000;

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                string responseText = reader.ReadToEnd();
                reader.Close();
                return responseText;
            } catch (Exception) {
                return "";
            }
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="para"></param>
        public static string HttpPost(string url, string para, string contentType) {
            try {
                HttpWebRequest request = (System.Net.HttpWebRequest)WebRequest.Create(url);
                request.ContentType = contentType;
                request.Method = "post";
                request.Timeout = 6000;

                Encoding encoding = Encoding.GetEncoding("UTF-8");
                byte[] data = encoding.GetBytes(para);
                Stream stream = request.GetRequestStream();
                stream.Write(data, 0, data.Length);
                stream.Flush();
                stream.Close();

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                string responseText = reader.ReadToEnd();
                reader.Close();
                return responseText;
            } catch (Exception) {
                return "";
            }
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string HttpPost(string url, Hashtable table) {
            string contenttype = "application/json";
            //发送请求 
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            JavaScriptSerializer se = new JavaScriptSerializer();
            string para = SerializationProvider.ToJson(table);
            request.Method = "post";
            request.ContentType = contenttype;
            request.ContentLength = para.Length;
            Stream stream = request.GetRequestStream();

            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] data = encoding.GetBytes(para);
            stream.Write(data, 0, para.Length);
            stream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            string responseText = reader.ReadToEnd();
            reader.Close();
            return responseText;
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string HttpPost(string url, string auth, Hashtable table) {
            string contenttype = "application/json";
            //发送请求 
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            JavaScriptSerializer se = new JavaScriptSerializer();
            string para = se.Serialize(table);
            request.Method = "post";
            request.ContentType = contenttype;
            request.Accept = "application/json";
            request.ContentLength = para.Length;
            request.Headers["Authorization"] = auth;
            Stream stream = request.GetRequestStream();

            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] data = encoding.GetBytes(para);
            stream.Write(data, 0, para.Length);
            stream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            string responseText = reader.ReadToEnd();
            reader.Close();
            return responseText;
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string InspireHttpPost(string url, string majorDomain, string subDomain, string id,
            string accessKey, string timestamp, string timeout, string nonce, string sign, Hashtable json, StringBuilder sb) {
            string contenttype = "text/json";
            //发送请
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            JavaScriptSerializer se = new JavaScriptSerializer();

            string para = SerializationProvider.ToJson(json);
            sb.AppendLine("InspireHttpPost=>" + para);
            request.Method = "post";
            request.ContentType = contenttype;
            request.Headers["X-Zc-Major-Domain"] = majorDomain;
            request.Headers["X-Zc-Major-Domain"] = majorDomain;
            request.Headers["X-Zc-Sub-Domain"] = subDomain;
            request.Headers["X-Zc-Developer-Id"] = id;
            request.Headers["X-Zc-Access-Key"] = accessKey;
            request.Headers["X-Zc-Timestamp"] = timestamp;
            request.Headers["X-Zc-Timeout"] = timeout;
            request.Headers["X-Zc-Nonce"] = nonce;
            request.Headers["X-Zc-Developer-Signature"] = sign;
            Stream stream = request.GetRequestStream();
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] data = encoding.GetBytes(para);
            stream.Write(data, 0, data.Length);
            stream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            reader.Close();
            return response.Headers["X-Zc-Msg-Name"];
        }

        /// <summary>
        /// http/https请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string CreateHttpPost(string url, Hashtable table) {
            HttpWebRequest request = null;
            //如果是发送HTTPS请求  
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase)) {
                request = (HttpWebRequest)WebRequest.Create(url);
                X509Certificate2 cert = CreateX509Certificate2();
                request.ClientCertificates.Add(cert);
            } else {
                //发送请求 
                request = (HttpWebRequest)WebRequest.Create(url);
            }
            string para = SerializationProvider.ToJson(table);
            request.Method = "post";
            request.ContentType = "application/json";
            Stream stream = request.GetRequestStream();

            byte[] data = Encoding.UTF8.GetBytes(para);
            stream.Write(data, 0, data.Length);
            stream.Flush();
            stream.Close();
            return CreateHttpResponse(request);
        }

        /// <summary>  
        /// 创建HttpResponse  
        /// </summary>  
        /// <param name="request"></param>  
        /// <returns></returns>  
        public static String CreateHttpResponse(HttpWebRequest request) {
            String str;
            HttpWebResponse response = null;
            Stream responseStream = null;
            StreamReader responseReader = null;
            try {
                using (response = (HttpWebResponse)request.GetResponse()) {
                    responseStream = response.GetResponseStream();
                    responseReader = new StreamReader(responseStream, Encoding.UTF8);
                    StringBuilder sb = new StringBuilder();
                    sb.Append(responseReader.ReadToEnd());
                    str = sb.ToString();
                }
            } catch (Exception e) {
                str = "{\"rescode\":\"0\",\"resmsg\":\"通信失败。原因：" + e.Message + "\"}";
            } finally {
                if (null != response) {
                    responseReader.Close();
                    responseStream.Close();
                    response.Close();
                }
            }
            return str;
        }

        /// <summary>  
        /// 创建X509证书  
        /// </summary>  
        /// <returns></returns>  
        public static X509Certificate2 CreateX509Certificate2() {
            X509Certificate2 cert = null;
            try {
                cert = new X509Certificate2(certFilePath, certFilePwd);
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(ServerCertificateValidationCallback);
            } catch (Exception e) {
                Console.WriteLine("创建X509Certificate2失败。原因：" + e.Message);
                cert = null;
            }
            return cert;
        }

        private static bool ServerCertificateValidationCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) {
            return true; //总是接受  
        }

    }
}
