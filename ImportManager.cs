namespace Storage.Services.ExportImport
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Web;
    using OfficeOpenXml;
    using System.Collections.Generic;
    using Storage.Data.Models;

    /// <summary>
    /// Import manager
    /// </summary>
    public partial class ImportManager : IImportManager
    {

        #region Utilities

        protected virtual int GetColumnIndex(string[] properties, string columnName)
        {
            if (properties == null)
                throw new ArgumentNullException("properties");

            if (columnName == null)
                throw new ArgumentNullException("columnName");

            for (int i = 0; i < properties.Length; i++)
                if (properties[i].Equals(columnName, StringComparison.InvariantCultureIgnoreCase))
                    return i + 1; //excel indexes start from 1
            return 0;
        }

        protected virtual string ConvertColumnToString(object columnValue)
        {
            if (columnValue == null)
                return null;

            return Convert.ToString(columnValue);
        }

        protected virtual string GetMimeTypeFromFilePath(string filePath)
        {
            var mimeType = MimeMapping.GetMimeMapping(filePath);

            //little hack here because MimeMapping does not contain all mappings (e.g. PNG)
            if (mimeType == "application/octet-stream")
                mimeType = "image/jpeg";

            return mimeType;
        }
        #endregion

        #region Method

        /// <summary>
        /// Import HAWB data from XLSX file
        /// </summary>
        /// <param name="stream">Stream</param>
        public IList<HawbImportInfo> ImportHAWBFromXlsx(Stream stream)
        {
            var list = new List<HawbImportInfo>();
            try
            {
                // ok, we can run the real code of the sample now
                using (var xlPackage = new ExcelPackage(stream))
                {
                    // get the first worksheet in the workbook
                    var worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                        throw new Exception("No worksheet found");

                    //the columns
                    var properties = new string[]
                                    {
                                        "Shipper",
                                        "Receiving Agent",
                                        "Estimated Time of Arrival",
                                        "Voyage",
                                        "Consignee",
                                        "Master Bill",
                                        "House Bill Number",
                                        "Packs",
                                        "Type",
                                        "Weight",
                                        "UW",
                                        "20GP",
                                        "20RE",
                                        "40GP",
                                        "40RE",
                                        "ACI Msg. Status",
                                        "ACI Status",
                                        "Chargeable",
                                        "Unit",
                                        "Act. Delivery",
                                        "Actual Pickup",
                                        "Volume",
                                        "UV"
                                    };

                    int iRow = 2;
                    while (true)
                    {
                        bool allColumnsAreEmpty = true;
                        for (var i = 1; i <= properties.Length; i++)
                        {
                            if (worksheet.Cells[iRow, i].Value != null && !String.IsNullOrEmpty(worksheet.Cells[iRow, i].Value.ToString()))
                            {
                                allColumnsAreEmpty = false;
                                break;
                            }
                        }

                        if (allColumnsAreEmpty)
                            break;

                        var shipper = worksheet.Cells[iRow, GetColumnIndex(properties, "Shipper")].Text;
                        var forwarder = worksheet.Cells[iRow, GetColumnIndex(properties, "Receiving Agent")].Text;
                        var arriveDate = GetFormatDate(worksheet.Cells[iRow, GetColumnIndex(properties, "Estimated Time of Arrival")].Text);
                        var voyage = worksheet.Cells[iRow, GetColumnIndex(properties, "Voyage")].Text;
                        var consignee = worksheet.Cells[iRow, GetColumnIndex(properties, "Consignee")].Text;
                        var masterBill = worksheet.Cells[iRow, GetColumnIndex(properties, "Master Bill")].Text;
                        var hAWBNo = worksheet.Cells[iRow, GetColumnIndex(properties, "House Bill Number")].Text;
                        var packages = GetFormatInt(worksheet.Cells[iRow, GetColumnIndex(properties, "Packs")].Text);
                        var packageType = worksheet.Cells[iRow, GetColumnIndex(properties, "Type")].Text;
                        var grossWeigth = GetFormatDecimal(worksheet.Cells[iRow, GetColumnIndex(properties, "Weight")].Text);
                        var uw = worksheet.Cells[iRow, GetColumnIndex(properties, "UW")].Text;
                        var chargeable = GetFormatDecimal(worksheet.Cells[iRow, GetColumnIndex(properties, "Chargeable")].Text);
                        var unit = worksheet.Cells[iRow, GetColumnIndex(properties, "Unit")].Text;
                        var volume = GetFormatDecimal(worksheet.Cells[iRow, GetColumnIndex(properties, "Volume")].Text);
                        var uv = worksheet.Cells[iRow, GetColumnIndex(properties, "UV")].Text;

                        var entity = new HawbImportInfo()
                        {
                            Shipper = shipper.Trim(),
                            Forwarder = forwarder.Trim(),
                            ArrivedDate = arriveDate,
                            Voyage = voyage.Trim(),
                            Consignee = consignee.Trim(),
                            MasterBill = masterBill.Trim(),
                            HAWBNo = hAWBNo.Trim(),
                            Packages = packages,
                            OriginalPackageType = packageType.Trim(),
                            GrossWeigth = grossWeigth,
                            UW = uw.Trim(),
                            Chargeable = chargeable,
                            Unit = unit.Trim(),
                            Volume = volume,
                            UV = uv.Trim()
                        };


                        list.Add(entity);

                        //next product
                        iRow++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return list;
        }



        /// <summary>
        /// format String to Int, if format failed return null
        /// </summary>
        private int? GetFormatInt(string value)
        {
            int tempData;
            if (int.TryParse(value, out tempData))
            {
                return tempData;
            }

            return null;
        }

        /// <summary>
        /// format String to Decimal, if format failed return null
        /// </summary>
        private decimal? GetFormatDecimal(string value)
        {
            decimal tempData;
            if (decimal.TryParse(value, out tempData))
            {
                return tempData;
            }

            return null;
        }

        /// <summary>
        /// format String to DateTime, if format failed return null
        /// </summary>
        public DateTime? GetFormatDate(string value)
        {

            DateTime tempData;
            if (DateTime.TryParse(value, out tempData))
            {
                return tempData;
            }

            return null;
        }
        #endregion

    }
}
