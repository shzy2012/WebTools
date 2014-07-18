namespace VAU.Web.CommonCode
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Web;

    /// <summary>
    /// Summary description for UploadFiles
    /// </summary>
    public class UploadFiles
    {
        #region Constructors and Destructors

        public UploadFiles()
        {
        }

        #endregion

        #region Public Properties

        public virtual string BillNo { get; set; }

        public virtual string FileExtension { get; set; }

        public virtual string FileName { get; set; }

        public virtual string FilePath { get; set; }

        public virtual decimal? FileSize { get; set; }

        public virtual string OrderType { get; set; }

        public virtual HttpPostedFile PostFile { get; set; }

        public virtual Guid Unique { get; set; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// Upload all file of current page
        /// </summary>
        /// <param name="httpFiles"> this.Request.Files </param>
        /// <param name="savePath"> </param>
        public void UploadFile(HttpFileCollection httpFiles, string savePath)
        {
            if (httpFiles == null || httpFiles.Count <= 0)
            {
                return;
            }

            try
            {
                var getKeys = httpFiles.AllKeys.Distinct();
                foreach (var key in getKeys)
                {
                    var files = httpFiles.GetMultiple(key);
                    foreach (HttpPostedFile item in files)
                    {
                        var path = savePath + item.FileName;
                        if (!File.Exists(path))
                        {
                            File.Delete(path);
                        }

                        item.SaveAs(path);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion
    }
}
