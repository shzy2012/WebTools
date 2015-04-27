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

    #endregion
    public class ExportManager
    {
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

        public static void ExportPaymnetToCSV(System.Web.HttpResponse response, IList<Bank> banks, IList<PaySetupPaymentHeader> payHead, IList<PaySetupPaymentLine> payLine)
        {
            byte[] bytes = new byte[0];
            using (var stream = new MemoryStream())
            {
                using (TextWriter writer = new StreamWriter(stream))
                {
                    #region MyRegion

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

                    var suppliers = payHead.Select(x => new { x.SupplierId }).Distinct();
                    foreach (var supllier in suppliers)
                    {
                        #region Head

                        var bank = banks.FirstOrDefault(x => x.SupplierId == supllier.SupplierId);
                        var tmpHead = (string[])columns.Clone();
                        var curDate = DateTime.Now.ToString("dd/MM/yyyy");
                        tmpHead[0] = "P";
                        tmpHead[1] = "TT";
                        tmpHead[2] = "ON";
                        tmpHead[6] = "HK";
                        tmpHead[7] = "HKG";
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
                        tmpHead[38] = curHeads.Sum(x => x.FinalPaidAmount).GetValueOrDefault().ToString();
                        dataList.Add(tmpHead);

                        foreach (var head in curHeads)
                        {
                            var tmpLine = (string[])columns.Clone();
                            tmpLine[0] = "I";

                            if (head.OrderType == "PO")
                            {
                                tmpLine[1] = "DEPOSIT";
                            }
                            else if (head.OrderType == "Prepay")
                            {
                                tmpLine[1] = "PREPAY";
                            }
                            else
                            {
                                tmpLine[1] = "INVOICE";
                            }

                            tmpLine[2] = curDate;
                            tmpLine[3] = head.BillNum.Replace(',', '_').Replace('\'', '_').Replace('"', '_'); //// Deal  , 92:' "
                            tmpLine[4] = head.FinalPaidAmount.GetValueOrDefault().ToString();
                            dataList.Add(tmpLine);
                        }

                        #endregion
                    }

                    var lastRow = (string[])columns.Clone();
                    lastRow[0] = "T";
                    lastRow[1] = suppliers.Count().ToString();
                    lastRow[2] = payHead.Sum(x => x.FinalPaidAmount.GetValueOrDefault()).ToString();
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
            response.AddHeader("Content-Disposition", string.Format("attachment;filename=Payment {0}.csv", DateTime.Now.ToString("yyyy-MM-dd")));

            response.BinaryWrite(bytes);
            if (response.IsClientConnected)
            {
                response.Flush();
            }

            response.End();
        }

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
    }
}
