namespace VAU.Web.CommonCode
{
    using System;
    using System.IO;
    using System.Web;

    /// <summary>
    /// Summary description for DownloadFile
    /// </summary>
    public class DownloadFile
    {
        public DownloadFile()
        {
        }

        /// <summary>
        /// Through the path to the file to download
        /// </summary>
        /// <param name="response">This page response</param>
        /// <param name="path">File path</param>
        public static void Download(HttpResponse response, string path, string filename = null)
        {
            try
            {
                var pushbyte = File.ReadAllBytes(path);
                FileInfo fileInfo = new FileInfo(path);
                response.Clear();
                response.ClearHeaders();
                response.ClearContent();
                response.ContentType = PageContentType.GetConteneType(fileInfo.Extension); // file type
                response.AddHeader("Content-Length", pushbyte.Length.ToString());
                if (string.IsNullOrWhiteSpace(filename))
                {
                    filename = fileInfo.Name;
                }

                response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(filename, System.Text.Encoding.UTF8).Replace("+", "%20"));

                response.BinaryWrite(pushbyte);
                if (response.IsClientConnected)
                {
                    response.Flush();
                }
            }
            catch (Exception ex)
            {
                PageBase.Log.Error("System log : " + ex);
                throw ex;
            }
        }
    }
}
