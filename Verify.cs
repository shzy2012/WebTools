namespace VAU.Web.CommonCode
{
    #region NameSpace

    using System;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;
    using System.Web;
    using System.Web.UI;
    using System.Web.UI.HtmlControls;

    using Spring.Context.Support;

    using VAU.Domain;
    using VAU.Security;
    using VAU.Security.Entity;

    #endregion

    /// <summary>
    /// All page inherit from PageBase
    /// </summary>
    public class PageBase : Spring.Web.UI.Page
    {

        /// <summary>
        /// format String to DateTime, if format failed return null
        /// </summary>
        public DateTime? GetFormatDate(string textBox)
        {
            DateTime tempData;
            if (DateTime.TryParse(textBox, out tempData))
            {
                return tempData;
            }

            return null;
        }

        /// <summary>
        /// format String to Decimal, if format failed return null
        /// </summary>
        public decimal? GetFormatDecimal(string textBox)
        {
            decimal tempData;
            if (decimal.TryParse(textBox, out tempData))
            {
                return tempData;
            }

            return null;
        }

        /// <summary>
        /// format String to Int, if format failed return null
        /// </summary>
        public int? GetFormatInt(string textBox)
        {
            int tempData;
            if (int.TryParse(textBox, out tempData))
            {
                return tempData;
            }

            return null;
        }

        /// <summary>
        /// format String to Int, if format failed return null
        /// </summary>
        public double? GetFormatDouble(string textBox)
        {
            double tempData;
            if (double.TryParse(textBox, out tempData))
            {
                return tempData;
            }

            return null;
        }

        public bool GetFormatBool(object obj)
        {
            if (obj == null)
            {
                return false;
            }

            bool outRst = false;
            if (bool.TryParse(obj.ToString(), out outRst))
            {
                return outRst;
            }

            return outRst;
        }

        /// <summary>
        /// format String to Percent(like:00.00%), if format failed return string.Empty
        /// </summary>
        public string ConvertPercent(object textBox)
        {
            return string.Format("{0:0.00%}", textBox);
        }

        /// <summary>
        /// get the IP Address
        /// </summary>
        /// <returns> </returns>
        public string GetIp()
        {
            var result = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(result))
            {
                result = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }

            if (string.IsNullOrEmpty(result))
            {
                result = HttpContext.Current.Request.UserHostAddress;
            }

            if (string.IsNullOrEmpty(result) || result.Contains("::1"))
            {
                return "127.0.0.1";
            }

            return result;
        }

        /// <summary>
        /// get current user's PC information
        /// </summary>
        /// <returns> </returns>
        public string GetMachineInfo()
        {
            string curInfo;
            try
            {
                ////var browser = Request.Browser;
                ////curInfo = "Browser:" + browser.Browser + ", Ver:" + browser.Version + ", Platform:" + browser.Platform;
                curInfo = this.Request.UserAgent;
            }
            catch (Exception)
            {
                return string.Empty;
            }

            return curInfo;
        }

        /// <summary>
        /// Alert a message on client window
        /// </summary>
        /// <param name="alertstrs"> </param>
        public void LoadAlert(string alertstrs)
        {
            var msg = this.Form.Parent.FindControl("alertMsg") as HtmlContainerControl;
            if (msg != null)
            {
                msg.InnerHtml = alertstrs;
            }

            var csname = "ButtonClickScriptX1";
            Type cstype = this.GetType();
            ClientScriptManager cs = this.Page.ClientScript;
            if (!cs.IsClientScriptBlockRegistered(cstype, csname))
            {
                cs.RegisterStartupScript(cstype, csname, string.Format("<script>ShowAlert();</script>", string.Empty));
            }
        }

        /// <summary>
        /// get whether string contains a key in keys
        /// </summary>
        /// <param name="strs"> </param>
        /// <param name="key"> </param>
        /// <returns> </returns>
        public bool StringContain(string strs, string[] key)
        {
            return key.Any(k => strs != null && strs.Contains(k));
        }

        #endregion

        #region Methods

        /// <summary>
        /// if have error in the page , use log4net write the error
        /// </summary>
        public void Page_Error(object sender, EventArgs args)
        {
            // Get the latest error
            var ex = this.Server.GetLastError();

            // write the error log
            Log.Error(ex.Message + ex.StackTrace, ex);

            // clear the info of log
            this.Server.ClearError();

            // if error from vauorderlist.aspx or vauorderdetail.aspx, the file path is different
            var path = this.Request.ApplicationPath;
            this.Response.Redirect(
                sender.ToString().ToLower().IndexOf("suppliers") != -1 ? "~/ErrorPage.aspx" : "~/ErrorPage.aspx?GoTo=1");
        }

        /// <summary>
        /// call the page permission method
        /// </summary>
        public void Page_PreInit(object sender, EventArgs e)
        {
            var strArray = sender.ToString().Split('_');
            if (strArray.Length == 4)
            {
                if (!this.Authorization.ValidateFunctionResource(strArray[2]))
                {
                    this.Response.Redirect("~/Pages/System/WebNavigation.aspx?validate=0");
                }
            }
        }

        /// <summary>
        /// call the button permission method
        /// </summary>
        /// <param name="btnId"> </param>
        /// <returns> </returns>
        public bool ValidateButtonResource(string btnId)
        {
            return this.Authorization.ValidateButtonResource(btnId);
        }

        public void SaveAs(HttpPostedFile postedFile, string bomSavePath)
        {
            try
            {
                if (!Directory.Exists(bomSavePath))
                {
                    Directory.CreateDirectory(bomSavePath);
                }

                var file = bomSavePath + postedFile.FileName;
                if (File.Exists(file))
                {
                    File.Delete(file);
                }

                postedFile.SaveAs(file);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool VerifyExtension(string extenName)
        {
            string[] extension = { "xlsx", "xls", "doc", "docx", "pdf", "jpg", "png", "gif", "bmp", "jpeg" };
            if (extension.Contains(extenName.ToLower()))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// </summary>
        /// <param name="patten"> Regex patten </param>
        /// <param name="verifyStr"> verification string </param>
        /// <returns> find count </returns>
        public int Regex(string patten, string verifyStr)
        {
            if (verifyStr == null)
            {
                return 0;
            }

            Regex rgx = new Regex(patten, RegexOptions.IgnoreCase);
            MatchCollection matches = rgx.Matches(verifyStr);
            return matches.Count;
        }

        public bool FileExtension(object path)
        {
            if (path == null)
            {
                return false;
            }

            var extName = Path.GetExtension(path.ToString()).ToLower();
            return extName == ".pdf";
        }

        #endregion
    }
}
