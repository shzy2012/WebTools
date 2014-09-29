namespace VAU.Web.CommonCode
{
    #region using

    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Web;

    using OfficeOpenXml;
    using OfficeOpenXml.Style;

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
                worksheet.Cells["A1:O1"].AutoFilter = true;

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
