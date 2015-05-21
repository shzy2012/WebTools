namespace VAU.Web.CommonCode
{
    #region using

    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Web;

    using OfficeOpenXml;
    using OfficeOpenXml.Style;

    using VAU.Domain;
    using VAU.Dto;
    using VAU.EnumType;

    using Excel = Microsoft.Office.Interop.Excel;
    #endregion
    public class ExportExcelManager
    {
        #region Business Export with style

        /// <summary>
        /// Using like,
        /// var data = new List<BOM>; 
        /// ExportExcelManager.ExportBomToXlsx(this.Response, data);
        /// </summary>
        /// <param name="response"></param>
        /// <param name="bom"></param>
        public static void ExportBomToXlsx(System.Web.HttpResponse response, IList<BOM4Excel> bom)
        {
            var stream = new MemoryStream();

            // ok, we can run the real code of the sample now
            using (var xlpackage = new ExcelPackage(stream))
            {
                //// uncomment this line if you want the XML written out to the outputDir
                //// xlPackage.DebugMode = true; 
                //// get handle to the existing worksheet
                var worksheet = xlpackage.Workbook.Worksheets.Add("BOMList");

                // Create Headers and format them
                var properties = new string[]
                    {
                        //// order properties
                        "GWMasterSKU",
                        "GKMasterSKU",
                        "ProductSpecialist",
                        "USAManager",
                        "CustomerMaster",
                        "SupplierName",
                        "SKUCount",
                        "ReportYear",
                        "BOMType",
                        "Status",
                        "Created",
                        "StatusUpdateDate",
                        "Interval1",
                        "SampleApproveDate",
                        "Interval2",
                        "ActiveStatus"
                    };

                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i];
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 32, 96));
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(1).Height = 30.00D;
                worksheet.Cells["A1:P1"].AutoFilter = true;

                for (int i = 1; i < properties.Length; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                int row = 2;
                foreach (var order in bom)
                {
                    int col = 1;

                    // order properties
                    worksheet.Cells[row, col].Value = order.GWMasterSKU;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.GKMasterSKU;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.ProductSpecialist;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.USAManager;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.CustomerMaster;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.SupplierName;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.SKUCount;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.ReportYear;
                    col++;

                    worksheet.Cells[row, col].Value = order.BOMType;
                    col++;

                    worksheet.Cells[row, col].Value = order.Status;
                    col++;

                    worksheet.Cells[row, col].Value = order.Created;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "yyyy/mm/dd HH:mm";
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.StatusUpdateDate;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "yyyy/mm/dd HH:mm";
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.Interval1;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.SampleApproveDate;
                    worksheet.Cells[row, col].Style.Numberformat.Format = "yyyy/mm/dd HH:mm";
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.Interval2;
                    worksheet.Cells[row, col].Style.ShrinkToFit = true;
                    col++;

                    worksheet.Cells[row, col].Value = order.ActiveStatus;
                    col++;

                    // next row
                    row++;
                }

                worksheet.View.FreezePanes(2, 1);

                xlpackage.Save();
            }

            byte[] bytes = stream.ToArray();

            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode("BOM.xlsx", System.Text.Encoding.UTF8).Replace("+", "%20"));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="response"></param>
        /// <param name="banks"></param>
        /// <param name="payHead"></param>
        public static void ExportPaymnetToCSV(System.Web.HttpResponse response, IList<Bank> banks, IList<PaySetupPaymentHeader> payHead)
        {
            byte[] bytes = new byte[0];
            using (var stream = new MemoryStream())
            {
                using (TextWriter writer = new StreamWriter(stream))
                {
                    #region Creat Colomn

                    // 62 
                    var columns = new string[]
                    {
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty,
                          string.Empty
                    };
                    #endregion

                    List<string[]> dataList = new List<string[]>();
                    var firRow = (string[])columns.Clone();
                    firRow[0] = "H";
                    firRow[1] = "P";
                    dataList.Add(firRow);

                    var suppliers = payHead.Select(x => new { x.SupplierId, x.BankId }).Distinct();
                    foreach (var supllier in suppliers)
                    {
                        #region Head

                        var bank = banks.FirstOrDefault(x => x.Id == supllier.BankId);
                        if (bank == null)
                        {
                            throw new Exception("No found the bank infomation");
                        }

                        if (string.IsNullOrWhiteSpace(bank.PaymentType))
                        {
                            throw new Exception("No found PaymentType of the bank");
                        }

                        var tmpHead = (string[])columns.Clone();
                        var curDate = DateTime.Now.ToString("dd/MM/yyyy");
                        tmpHead[0] = "P";
                        tmpHead[1] = bank.PaymentType;
                        tmpHead[2] = "ON";
                        tmpHead[6] = "HK";
                        tmpHead[7] = "HKG";
                        tmpHead[8] = "44717855891";
                        tmpHead[9] = curDate;
                        tmpHead[22] = "0";
                        tmpHead[29] = "0";
                        tmpHead[30] = "0";
                        tmpHead[33] = "0";
                        tmpHead[34] = "0";
                        tmpHead[35] = "0";
                        tmpHead[36] = "4";
                        tmpHead[37] = "USD";
                        tmpHead[39] = "C";
                        tmpHead[40] = "P";
                        tmpHead[48] = "O";
                        tmpHead[61] = bank.Id.ToString();
                        #endregion

                        #region Line

                        var curHeads = payHead.Where(x => x.SupplierId == supllier.SupplierId);
                        tmpHead[20] = curHeads.First().SNReference;
                        tmpHead[38] = curHeads.Sum(x => x.ToPayAll).GetValueOrDefault().ToString();
                        dataList.Add(tmpHead);

                        foreach (var head in curHeads)
                        {
                            var tmpLine = (string[])columns.Clone();
                            tmpLine[0] = "I";

                            if (head.PaymentType == "PO")
                            {
                                tmpLine[1] = "DEPOSIT";
                            }
                            else if (head.PaymentType == "Prepay")
                            {
                                tmpLine[1] = "PREPAY";
                            }
                            else
                            {
                                tmpLine[1] = "INVOICE";
                            }

                            tmpLine[2] = curDate;
                            tmpLine[3] = head.BillNum.Replace(',', '_').Replace('\'', '_').Replace('"', '_'); //// Deal  , 92:' "
                            tmpLine[4] = head.ToPayAll.GetValueOrDefault().ToString();
                            dataList.Add(tmpLine);
                        }

                        #endregion
                    }

                    var lastRow = (string[])columns.Clone();
                    lastRow[0] = "T";
                    lastRow[1] = suppliers.Count().ToString();
                    lastRow[2] = payHead.Sum(x => x.ToPayAll.GetValueOrDefault()).ToString();
                    dataList.Add(lastRow);

                    foreach (var item in dataList)
                    {
                        writer.WriteLine(string.Join(",", item));
                    }
                }

                bytes = stream.ToArray();
            }

            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.ContentType = "text/csv";
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.AddHeader("Content-Disposition", string.Format("attachment;filename=Payment{0}.csv", DateTime.Now.ToString("yyyy-MM-dd")));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }

            response.End();
        }

        public static void ExportPurasingReportToExcel<T>(System.Web.HttpResponse response, IList<T> list)
        {
            var properties = new List<string>();

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in props)
            {
                properties.Add(prop.Name);
            }

            #region Subtitles

            var subtitles = new List<string>();
            subtitles.Add("SUPPLIER ID");
            subtitles.Add("FID");
            subtitles.Add("KINGDEE VENDOR SHORT NAME");
            subtitles.Add("KINGDE VENDOR FULL NAME");
            subtitles.Add("PS");
            subtitles.Add("OPEN ORDER VALUE - $$");
            subtitles.Add("OPEN ORDER VALUE - QTY");
            subtitles.Add("YTD INVOICED - $$");
            subtitles.Add("YTD INVOICED - QTY");
            subtitles.Add("PY INVOICED - $$");
            subtitles.Add("PY INVOICED – QTY");
            subtitles.Add("MTD - $$");
            subtitles.Add("MTD  - QTY");
            subtitles.Add("PREV MONTH - 01 - $$");
            subtitles.Add("PREV MONTH - 01 - QTY");
            subtitles.Add("PREV MONTH - 02 - $$");
            subtitles.Add("PREV MONTH - 02-QTY");
            subtitles.Add("PREV MONTH - 03 - $$");
            subtitles.Add("PREV MONTH - 03-QTY");
            subtitles.Add("PREV MONTH - 04 - $$");
            subtitles.Add("PREV MONTH - 04-QTY");
            subtitles.Add("PREV MONTH - 05 - $$");
            subtitles.Add("PREV MONTH - 05-QTY");
            subtitles.Add("PREV MONTH - 06 - $$");
            subtitles.Add("PREV MONTH - 06-QTY");
            subtitles.Add("PREV MONTH - 07 - $$");
            subtitles.Add("PREV MONTH - 07-QTY");
            subtitles.Add("PREV MONTH - 08 - $$");
            subtitles.Add("PREV MONTH - 08-QTY");
            subtitles.Add("PREV MONTH - 09 - $$");
            subtitles.Add("PREV MONTH - 09-QTY");
            subtitles.Add("PREV MONTH - 10 - $$");
            subtitles.Add("PREV MONTH - 10-QTY");
            subtitles.Add("PREV MONTH - 11 - $$");
            subtitles.Add("PREV MONTH - 11-QTY");
            subtitles.Add("PREV MONTH - 12 - $$");
            subtitles.Add("PREV MONTH - 12-QTY");
            subtitles.Add("TOTAL OPEN PRE - PAYMENTS");
            subtitles.Add("TOTAL OPEN UN-APPROVED ORDERS");
            subtitles.Add("TOTAL OPEN ORDERS WITHOUT DEPOSITS BEING PAID");
            subtitles.Add("YTD - TOTAL PAYMENTS MADE TO VENDOR");
            subtitles.Add("NOT DUE YET");
            subtitles.Add("X <= 15 ");
            subtitles.Add("15 < X <= 30");
            subtitles.Add("30 < X <= 45");
            subtitles.Add("45 < X <= 60");
            subtitles.Add("60 < X <= 75");
            subtitles.Add("75 < X <= 90");
            subtitles.Add("90 < X <= 120");
            subtitles.Add("120 + DAYS ");
            subtitles.Add("COMBINED AGING");
            #endregion

            var stream = new MemoryStream();
            using (var xlpackage = new ExcelPackage(stream))
            {
                var worksheet = xlpackage.Workbook.Worksheets.Add("VENDOR PURCHASING SMMRY");
                for (int i = 0; i < properties.Count; i++)
                {
                    worksheet.Cells[4, i + 1].Value = properties[i];
                    worksheet.Cells[4, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[4, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 32, 96));
                    worksheet.Cells[4, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[4, i + 1].Style.Font.Color.SetColor(Color.White);
                }

                for (int i = 0; i < subtitles.Count; i++)
                {
                    worksheet.Cells[5, i + 1].Value = subtitles[i];
                    worksheet.Cells[5, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[5, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 32, 96));
                    worksheet.Cells[5, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[5, i + 1].Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Row(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(4).Height = 30.00D;
                worksheet.Row(5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(5).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(5).Height = 50.00D;
                worksheet.View.FreezePanes(6, 6); //// Freeze 5 row, 5 colomn
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                //// caculate
                worksheet.Cells["F2"].Formula = string.Format("=SUBTOTAL(9,F{0}:F{1})", 6, list.Count + 5);
                worksheet.Cells["AL2"].Formula = string.Format("=SUBTOTAL(9,AL{0}:AL{1})", 6, list.Count + 5);
                worksheet.Cells["AM2"].Formula = string.Format("=SUBTOTAL(9,AM{0}:AM{1})", 6, list.Count + 5);
                worksheet.Cells["AN2"].Formula = string.Format("=SUBTOTAL(9,AN{0}:AN{1})", 6, list.Count + 5);
                worksheet.Cells["AP2"].Formula = string.Format("=SUBTOTAL(9,AP{0}:AP{1})", 6, list.Count + 5);
                worksheet.Cells["AQ2"].Formula = string.Format("=SUBTOTAL(9,AQ{0}:AQ{1})", 6, list.Count + 5);
                worksheet.Cells["AR2"].Formula = string.Format("=SUBTOTAL(9,AR{0}:AR{1})", 6, list.Count + 5);
                worksheet.Cells["AR2"].Formula = string.Format("=SUBTOTAL(9,AR{0}:AR{1})", 6, list.Count + 5);
                worksheet.Cells["AT2"].Formula = string.Format("=SUBTOTAL(9,AT{0}:AT{1})", 6, list.Count + 5);
                worksheet.Cells["AU2"].Formula = string.Format("=SUBTOTAL(9,AU{0}:AU{1})", 6, list.Count + 5);
                worksheet.Cells["AV2"].Formula = string.Format("=SUBTOTAL(9,AV{0}:AV{1})", 6, list.Count + 5);
                worksheet.Cells["AW2"].Formula = string.Format("=SUBTOTAL(9,AW{0}:AW{1})", 6, list.Count + 5);
                worksheet.Cells["AX2"].Formula = string.Format("=SUBTOTAL(9,AX{0}:AX{1})", 6, list.Count + 5);
                worksheet.Cells["AY2"].Formula = string.Format("=SUBTOTAL(9,AY{0}:AY{1})", 6, list.Count + 5);
                int row = 6;
                foreach (T item in list)
                {
                    int col = 1;

                    var values = new object[props.Length];
                    for (int i = 0; i < props.Length; i++)
                    {
                        values[i] = props[i].GetValue(item, null);
                        worksheet.Cells[row, col].Value = values[i];
                        if (props[i].PropertyType.GenericTypeArguments.Length > 0)
                        {
                            var currenttype = props[i].PropertyType.GenericTypeArguments[0].FullName;
                            if (currenttype == typeof(DateTime).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "yyyy/mm/dd HH:mm";
                            }

                            if (currenttype == typeof(decimal).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.0";
                            }

                            if (currenttype == typeof(double).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.0";
                            }

                            worksheet.Cells[row, 4].Style.ShrinkToFit = true;
                        }

                        col++;
                    }

                    // next row
                    row++;
                }

                xlpackage.Save();
            }

            byte[] bytes = stream.ToArray();

            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode("VENDOR PURCHASING SMMRY.xlsx", System.Text.Encoding.ASCII));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }
        }
        #endregion

        #region Export form excel file

        public static void RefreshExcel(string execelLocation)
        {
            object missingValue = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();
            excel.DisplayAlerts = false;
            Excel.Workbook theWorkbook = excel.Workbooks.Open(
                execelLocation,
                missingValue,
                false,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue,
                missingValue);

            lock (theWorkbook)
            {
                theWorkbook.RefreshAll();
            }

            System.Threading.Thread.Sleep(5 * 1000); // Make sure correct save 

            theWorkbook.Save();
            theWorkbook.Close();
            excel.Quit();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public static void ExportExcel(System.Web.HttpResponse response, string excelLocation)
        {
            byte[] bytes = File.ReadAllBytes(excelLocation);
            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode(Path.GetFileName(excelLocation), System.Text.Encoding.UTF8).Replace("+", "%20"));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }
        }

        #endregion

        #region Export excel from  data list, without style

        public static void ResponseExcel<T>(System.Web.HttpResponse response, List<T> items)
        {
            try
            {
                string attachment = "attachment; filename=vauExcel.xls";
                response.ClearContent();
                response.AddHeader("content-disposition", attachment);
                response.ContentType = "application/vnd.ms-excel";
                string tab = string.Empty;

                // Get all the properties
                PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in props)
                {
                    response.Write(tab + prop.Name);
                    tab = "\t";
                }

                response.Write("\n");
                foreach (T item in items)
                {
                    var values = new object[props.Length];
                    for (int i = 0; i < props.Length; i++)
                    {
                        values[i] = props[i].GetValue(item, null);
                        if (values[i] != null)
                        {
                            response.Write(values[i].ToString().Trim() + "\t");
                        }
                        else
                        {
                            response.Write("\t");
                        }
                    }

                    response.Write("\n");
                }

                response.Flush();
                response.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Template of excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="response"></param>
        /// <param name="list"></param>
        public static void ExportToExcel<T>(System.Web.HttpResponse response, IList<T> list)
        {
            var properties = new List<string>();

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in props)
            {
                properties.Add(prop.Name);
            }

            var stream = new MemoryStream();

            using (var xlpackage = new ExcelPackage(stream))
            {
                //// uncomment this line if you want the XML written out to the outputDir
                //// xlPackage.DebugMode = true; 
                //// get handle to the existing worksheet
                var worksheet = xlpackage.Workbook.Worksheets.Add("sheet");
                for (int i = 0; i < properties.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i];
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 32, 96));
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(Color.White);
                }

                worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Row(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Row(1).Height = 30.00D;

                int row = 2;
                foreach (T item in list)
                {
                    int col = 1;

                    var values = new object[props.Length];
                    for (int i = 0; i < props.Length; i++)
                    {
                        values[i] = props[i].GetValue(item, null);
                        worksheet.Cells[row, col].Value = values[i];
                        if (props[i].PropertyType.GenericTypeArguments.Length > 0)
                        {
                            var currenttype = props[i].PropertyType.GenericTypeArguments[0].FullName;
                            if (currenttype == typeof(DateTime).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "yyyy/mm/dd HH:mm";
                            }

                            if (currenttype == typeof(decimal).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.0";
                            }

                            if (currenttype == typeof(double).FullName)
                            {
                                worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0.0";
                            }

                            worksheet.Cells[row, col].Style.ShrinkToFit = true;
                        }

                        col++;
                    }

                    // next row
                    row++;
                }

                worksheet.View.FreezePanes(2, 1);
                xlpackage.Save();
            }

            byte[] bytes = stream.ToArray();

            response.Clear();
            response.ClearHeaders();
            response.ClearContent();
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.AddHeader("Content-Disposition", "attachment; filename=" + HttpUtility.UrlEncode("general.xlsx", System.Text.Encoding.UTF8).Replace("+", "%20"));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }
        }
        #endregion
    }
}
