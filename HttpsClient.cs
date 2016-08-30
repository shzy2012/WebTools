using System;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Net.Security;

namespace HttpsClient {
    class Program {
        static void Main(string[] args) {
            string url = "https://115.231.94.9:8007/api/watch/ActiveWatch";
            string contentType = "application/json";
            string data = "{}";
            string hGet = HttpsPost(url, data, contentType);
            Console.WriteLine(hGet);
            Console.ReadLine();
        }
        public bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors) {   // 总是接受  
            return true;
        }

        // callback used to validate the certificate in an SSL conversation
        private static bool ValidateRemoteCertificate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors policyErrors) {
            bool result = false;
            if(cert.Subject.ToUpper().Contains("CN=WINDOWS2012R2")) {
                result = true;
            }

            return result;
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="para"></param>
        public static string HttpsGet(string url, string para, string contentType) {
            try {
                url = url + "?" + para;
                //验证服务器证书回调自动验证
                ServicePointManager.ServerCertificateValidationCallback += new RemoteCertificateValidationCallback(ValidateRemoteCertificate);

                HttpWebRequest request = (System.Net.HttpWebRequest)WebRequest.Create(url);
                request.Proxy = null;
                request.Credentials = CredentialCache.DefaultCredentials;
                X509Certificate x509 = new X509Certificate("ZS.pfx", "passwd");
                request.ClientCertificates.Add(x509);
                request.ContentType = contentType;
                request.Method = "GET";
                request.Timeout = 15000;
                //验证服务器证书回调自动验证
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                string responseText = reader.ReadToEnd();
                reader.Close();
                return responseText;
            } catch(Exception ex) {
                return ex.Message;
            }
        }

        /// <summary>
        /// http请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="para"></param>
        public static string HttpsPost(string url, string para, string contentType) {
            try {
                //验证服务器证书回调自动验证
                ServicePointManager.ServerCertificateValidationCallback += new RemoteCertificateValidationCallback(ValidateRemoteCertificate);

                HttpWebRequest request = (System.Net.HttpWebRequest)WebRequest.Create(url);
                request.Proxy = null;
                request.Credentials = CredentialCache.DefaultCredentials;
                X509Certificate x509 = new X509Certificate("ZS.pfx", "passwd");
                request.ClientCertificates.Add(x509);
                request.ContentType = contentType;
                request.Method = "POST";
                request.Timeout = 15000;

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
            } catch(Exception ex) {
                return ex.Message;
            }
        }
    }
}
