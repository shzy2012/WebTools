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
        public static void Download(HttpResponse response, string path)
        {
            Stream istream = null;
            byte[] buffer = new byte[10000];
            int length;
            long dataToRead;

            try
            {
                FileInfo fileInfo = new FileInfo(path);
                var filename = fileInfo.Name;
                istream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
                dataToRead = istream.Length;
                response.Clear();
                response.ClearHeaders();
                response.ClearContent();
                response.ContentType = PageContentType.GetConteneType(fileInfo.Extension); // file type
                response.AddHeader("Content-Length", dataToRead.ToString());
                response.AddHeader(
                    "Content-Disposition",
                    "attachment; filename=" + HttpUtility.UrlEncode(filename, System.Text.Encoding.UTF8).Replace("+", "%20"));
                while (dataToRead > 0)
                {
                    if (response.IsClientConnected)
                    {
                        length = istream.Read(buffer, 0, 10000);
                        response.OutputStream.Write(buffer, 0, length);
                        response.Flush();
                        buffer = new byte[10000];
                        dataToRead = dataToRead - length;
                    }
                    else
                    {
                        dataToRead = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (istream != null)
                {
                    istream.Close();
                }
            }
        }
    }
}
