namespace VAU.Web.CommonCode
{
    using System;
    using System.Collections.Generic;

    using Ionic.Zip;

    /// <summary>
    /// Summary description for OutputZip
    /// </summary>
    public class OutputZip
    {
        #region Public Methods and Operators

        /// <summary>
        /// Zip files and Download
        /// </summary>
        /// <param name="response">Current HttpResponse</param>
        /// <param name="zipFileToCreate">Zip save path</param>
        /// <param name="files">zipfile</param>
        public static void ResponseZip(System.Web.HttpResponse response, string zipFileToCreate, List<string> files)
        {
            try
            {
                using (ZipFile zip = new ZipFile())
                {
                    foreach (string filename in files)
                    {
                        ZipEntry e = zip.AddFile(filename, string.Empty);
                        e.Comment = "Added by  VAU";
                    }

                    zip.Comment =
                        string.Format(
                            "This zip archive was created by the CreateZip example application on machine '{0}'",
                            System.Net.Dns.GetHostName());

                    zip.Save(zipFileToCreate);
                }

                response.Clear();
                response.AppendHeader("Content-Disposition", "attachment; filename=VAUFiles.zip");
                response.ContentType = "application/x-zip-compressed";
                response.WriteFile(zipFileToCreate);
                if (response.IsClientConnected)
                {
                    response.Flush();
                    response.Close();
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Zip files
        /// </summary>
        /// <param name="zipFileToCreate"></param>
        /// <param name="files"></param>
        public static void ZipFiles(string zipFileToCreate, List<string> files)
        {
            try
            {
                using (ZipFile zip = new ZipFile())
                {
                    foreach (string filename in files)
                    {
                        ZipEntry e = zip.AddFile(filename, string.Empty);
                        e.Comment = "Added by  VAU";
                    }

                    zip.Comment =
                        string.Format(
                            "This zip archive was created by the CreateZip example application on machine '{0}'",
                            System.Net.Dns.GetHostName());

                    zip.Save(zipFileToCreate);
                }
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }

        #endregion
    }
}
