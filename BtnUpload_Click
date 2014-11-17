protected void BtnUpload_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            int calcUploadSuccess = 0;
            if (this.ddlCategory.SelectedValue == "0")
            {
                this.message.InnerText = "Please Choose One Category 请选择图片类型. ";
                return;
            }

            int supplierId = GetSupplierId();
            //// 0: supplier is null 
            if (supplierId == 0)
            {
                this.message.InnerText = "No Supplier ID 不存在该供应商编号 .";
                return;
            }

            for (int i = 0; i < this.files.Controls.Count; i++)
            {
                if (this.files.Controls[i] is FileUpload)
                {
                    FileUpload tempFile = this.files.Controls[i] as FileUpload;
                    if (!tempFile.HasFile)
                    {
                        continue;
                    }
                    ////Get upload file expand name
                    string fileExt = System.IO.Path.GetExtension(tempFile.FileName).Trim();
                    ////accept image formats
                    if (
                        !(fileExt == ".jpg" || fileExt == ".png" || fileExt == ".gif" || fileExt == ".bmp"
                          || fileExt == ".jpeg"))
                    {
                        sb.Append(string.Format("Image of {0} is suffix illegal . 图片 {0} 后缀名不合法 .", tempFile.FileName));
                        continue;
                    }
                    ////Get upload image size
                    int size = tempFile.PostedFile.ContentLength / 1024 / 1024;
                    ////less or equal to 4M
                    if (size > 4)
                    {
                        sb.Append(
                            string.Format(
                                "Name of {0}  Image size should be less than to 4M 名称是：{0} 图片应该小于4M .",
                                tempFile.FileName));
                        continue;
                    }

                    //// Filling int format:01,02
                    string category = this.ddlCategory.SelectedValue.PadLeft(2, '0');
                    //// if not exist else Create Folder 
                    CreateSupplierFolder(uploadPath);
                    //// make upload image name
                    string imgName = supplierId + "_" + category + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff")
                                     + ".jpg";
                    //// make path
                    string fullName = uploadPath + imgName;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        tempFile.FileContent.CopyTo(ms);
                        byte[] getMem = ms.ToArray();
                        string imgStr = System.Text.Encoding.UTF8.GetString(getMem);
                        //// Is true  image?
                        if (!IsImage(imgStr))
                        {
                            sb.Append(string.Format("Name of {0} Image Illegal  名称 {0} 的图片非法. ", tempFile.FileName));
                            continue;
                        }
                    }

                    try
                    {
                        ////Upload image  to Server
                        tempFile.SaveAs(fullName);
                        ////save upload image data to DB
                        Save(supplierId, imgName);
                        calcUploadSuccess++;
                    }
                    catch (Exception)
                    {
                        sb.Append("Upload Failed  上传过程出现失败. ");
                    }
                }
            }

            if (sb.ToString().Trim() != string.Empty)
            {
                sb.Append(string.Format(" Success Upload {0},成功上传{0}张图片.", calcUploadSuccess));
                this.message.InnerText = sb.ToString();
            }
            else
            {
                this.message.InnerText = string.Format("Upload Success. 上传成功. Total:{0}.", calcUploadSuccess);
            }
        }
        
        
        /// <summary>
        ///  Determine whether the image contained unsafe
        /// </summary>
        /// <param name="strContent">physical image path</param>
        /// <returns>ture or false</returns>
        protected bool IsImage(string strContent)
        {
            string str =
                "href|request|script|.getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|.saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language=";
            foreach (string s in str.Split('|'))
            {
                if (strContent.IndexOf(s) != -1)
                {
                    return false;
                }
            }

            return true;
        }
