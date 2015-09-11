namespace VAU.Web.CommonCode
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using OfficeOpenXml;
    using VAU.Domain;
    using VAU.Dto;

    public class ImportManager
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static IList<ExportPaymentDto> ImportPaymentFromXlsx(MemoryStream stream)
        {
            var list = new List<ExportPaymentDto>();
            try
            {
                using (var xlpackage = new ExcelPackage(stream))
                {
                    //// get the first worksheet in the workbook
                    var worksheet = xlpackage.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        throw new Exception("No worksheet found");
                    }

                    //// the columns
                    var properties = new string[]
                                    {
                                        "Payment Reference",
                                        "Customer Reference",
                                        "Type",
                                        "Status",
                                        "Payee/BeneficiaryName",
                                        "Payment/Value Date",
                                        "Debit Date",
                                        "Payment Currency",
                                        "Total Amount (Payment Currency)",
                                        "Exchange Rate",
                                        "Debit Currency",
                                        "Total Amount (Debit Currency)",
                                        "Debit Account",
                                        "DR BCE",
                                        "DR ACE",
                                        "Internal Memo",
                                        "Purpose of Payment",
                                        "Destination Purpose of Payment",
                                        "Beneficiary Account Type",
                                        "Receiver ID",
                                        "Receiver ID Type",
                                        "Customer No",
                                        "Listed Company Code",
                                        "Product Type",
                                        "Beneficiary Account",
                                        "Import Reference",
                                        "Transaction/Related Reference",
                                        "BO Reference",
                                        "Processing Date",
                                        "Cheque Number",
                                        "Cheque Cleared Date",
                                        "Batch Reference",
                                        "Payment Detail 1",
                                        "Payment Detail 2",
                                        "Payment Detail Local language 1",
                                        "Payment Detail Local language 2",
                                        "Payee/Beneficiary Bank Code",
                                        "TT Beneficiary Bank Details",
                                        "Local Bank Clearing Code",
                                        "Branch Code	Creation Date",
                                        "On Behalf Of Type",
                                        "Credit Date",
                                        "Payee/Beneficiary Local Language2",
                                        "On Behalf Of Name",
                                        "On Behalf Of Account/Party Identifier",
                                        "On Behalf Of Address1",
                                        "On Behalf Of Address2",
                                        "On Behalf Of Address3"
                                    };

                    int row = 2;
                    while (true)
                    {
                        bool allColumnsAreEmpty = true;
                        for (var i = 1; i <= properties.Length; i++)
                        {
                            if (worksheet.Cells[row, i].Value != null && !string.IsNullOrEmpty(worksheet.Cells[row, i].Value.ToString()))
                            {
                                allColumnsAreEmpty = false;
                                break;
                            }
                        }

                        if (allColumnsAreEmpty)
                        {
                            break;
                        }

                        var entity = new ExportPaymentDto();
                        entity.PaymentNo = worksheet.Cells[row, GetColumnIndex(properties, "Payment Reference")].Value.ToString().Trim();
                        entity.BatchReference = worksheet.Cells[row, GetColumnIndex(properties, "Batch Reference")].Value.ToString().Trim();
                        entity.BeneficiaryName = worksheet.Cells[row, GetColumnIndex(properties, "Payee/BeneficiaryName")].Value.ToString().Trim();
                        entity.BOReference = worksheet.Cells[row, GetColumnIndex(properties, "BO Reference")].Value.ToString().Trim();
                        entity.CustomerReference = worksheet.Cells[row, GetColumnIndex(properties, "Customer Reference")].Value.ToString().Trim();
                        entity.PaymentDetail1 = worksheet.Cells[row, GetColumnIndex(properties, "Payment Detail 1")].Value.ToString().Trim();
                        entity.BeneficiaryAccount = worksheet.Cells[row, GetColumnIndex(properties, "Beneficiary Account")].Value.ToString().Trim();
                        list.Add(entity);

                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return list;
        }

        public static IList<FactoryInfo> ImportFactoryInfoFromXlsx(MemoryStream stream)
        {
            var list = new List<FactoryInfo>();
            try
            {
                using (var xlpackage = new ExcelPackage(stream))
                {
                    //// get the first worksheet in the workbook
                    var worksheet = xlpackage.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        throw new Exception("No worksheet found");
                    }

                    //// the columns
                    var properties = new string[]
                                    {
                                        "Id",
                                        "SupplierId",
                                        "DateOfAudit",
                                        "FactoryName",
                                        "TypeOfProductsManufactured",
                                        "YearEstablished",
                                        "Province",
                                        "City",
                                        "Town",
                                        "Address",
                                        "NumberOfBuildings",
                                        "Size",
                                        "Certifications1",
                                        "Certifications2",
                                        "Certifications3",
                                        "Certifications4",
                                        "Certifications5",
                                        "ProductionEmployees",
                                        "QcStaffEmployees",
                                        "AdminStaffEmployees",
                                        "QaManagersEmployees",
                                        "DesignDepartmentEmployees",
                                        "EngineersEmployees",
                                        "SecurityEmployees",
                                        "AssemblyLines",
                                        "CapacityPerMonth",
                                        "Customer1",
                                        "Customer2",
                                        "Customer3",
                                        "HRManagerName",
                                        "ProductionManagerName",
                                        "ChiefOfQualityName",
                                        "ChiefOfEngineerName",
                                        "IsAuditByThirdPartyTesting",
                                        "LastThirdPartyAuditDate",
                                        "ThirdPartyAuditPerYear",
                                        "IsBureauVeritasPerformedAudit",
                                        "FactoryCode",
                                        "Dormitories",
                                        "Canteen",
                                        "SecuirtyCameras",
                                        "SecuirtyGuards",
                                    };

                    int row = 2;
                    while (true)
                    {
                        bool allColumnsAreEmpty = true;
                        for (var i = 1; i <= properties.Length; i++)
                        {
                            if (worksheet.Cells[row, i].Value != null && !string.IsNullOrEmpty(worksheet.Cells[row, i].Value.ToString()))
                            {
                                allColumnsAreEmpty = false;
                                break;
                            }
                        }

                        if (allColumnsAreEmpty)
                        {
                            break;
                        }

                        var entity = new FactoryInfo();
                        entity.Id = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "Id")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "Id")].Value.ToString().Trim());
                        entity.FactoryCode = worksheet.Cells[row, GetColumnIndex(properties, "FactoryCode")].Value.ToString().Trim();
                        entity.SupplierId = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "SupplierId")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "SupplierId")].Value.ToString().Trim());
                        entity.FactoryName = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "FactoryName")].Value).ToString().Trim();
                        entity.TypeOfProductsManufactured = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "TypeOfProductsManufactured")].Value).ToString().Trim();
                        if (worksheet.Cells[row, GetColumnIndex(properties, "YearEstablished")].Value == null)
                        {
                            entity.YearEstablished = null;
                        }
                        else
                        {
                            entity.YearEstablished = Convert.ToDateTime(worksheet.Cells[row, GetColumnIndex(properties, "YearEstablished")].Value);
                        }
                        
                        entity.Province = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Province")].Value).ToString().Trim();
                        entity.City = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "City")].Value).ToString().Trim();
                        entity.Town = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Town")].Value).ToString().Trim();
                        entity.Address = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Address")].Value).ToString().Trim();
                        entity.NumberOfBuildings = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "NumberOfBuildings")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "NumberOfBuildings")].Value.ToString().Trim());
                        entity.Size = Convert.ToDecimal(worksheet.Cells[row, GetColumnIndex(properties, "Size")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "Size")].Value.ToString().Trim());
                        entity.Certifications1 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Certifications1")].Value).ToString().Trim();
                        entity.Certifications2 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Certifications2")].Value).ToString().Trim();
                        entity.Certifications3 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Certifications3")].Value).ToString().Trim();
                        entity.Certifications4 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Certifications4")].Value).ToString().Trim();
                        entity.Certifications5 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Certifications5")].Value).ToString().Trim();
                        entity.ProductionEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "ProductionEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "ProductionEmployees")].Value.ToString().Trim());
                        entity.QcStaffEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "QcStaffEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "QcStaffEmployees")].Value.ToString().Trim());
                        entity.AdminStaffEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "AdminStaffEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "AdminStaffEmployees")].Value.ToString().Trim());
                        entity.QaManagersEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "QaManagersEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "QaManagersEmployees")].Value.ToString().Trim());
                        entity.DesignDepartmentEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "DesignDepartmentEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "DesignDepartmentEmployees")].Value.ToString().Trim());
                        entity.EngineersEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "EngineersEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "EngineersEmployees")].Value.ToString().Trim());
                        entity.SecurityEmployees = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "SecurityEmployees")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "SecurityEmployees")].Value.ToString().Trim());
                        entity.AssemblyLines = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "AssemblyLines")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "AssemblyLines")].Value.ToString().Trim());
                        entity.CapacityPerMonth = Convert.ToDecimal(worksheet.Cells[row, GetColumnIndex(properties, "CapacityPerMonth")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "CapacityPerMonth")].Value.ToString().Trim());
                        entity.Customer1 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Customer1")].Value).ToString().Trim();
                        entity.Customer2 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Customer2")].Value).ToString().Trim();
                        entity.Customer3 = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "Customer3")].Value).ToString().Trim();
                        entity.HRManagerName = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "HRManagerName")].Value).ToString().Trim();
                        entity.ProductionManagerName = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "ProductionManagerName")].Value).ToString().Trim();
                        entity.ChiefOfQualityName = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "ChiefOfQualityName")].Value).ToString().Trim();
                        entity.ChiefOfEngineerName = GetValue(worksheet.Cells[row, GetColumnIndex(properties, "ChiefOfEngineerName")].Value).ToString().Trim();
                        entity.IsAuditByThirdPartyTesting = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "IsAuditByThirdPartyTesting")].Value).ToString().Trim());
                        if (worksheet.Cells[row, GetColumnIndex(properties, "LastThirdPartyAuditDate")].Value == null)
                        {
                            entity.LastThirdPartyAuditDate = null;
                        }
                        else
                        {
                            entity.LastThirdPartyAuditDate = Convert.ToDateTime(worksheet.Cells[row, GetColumnIndex(properties, "LastThirdPartyAuditDate")].Text);
                        }

                        entity.ThirdPartyAuditPerYear = Convert.ToInt32(worksheet.Cells[row, GetColumnIndex(properties, "ThirdPartyAuditPerYear")].Value == null ? "0" : worksheet.Cells[row, GetColumnIndex(properties, "ThirdPartyAuditPerYear")].Value.ToString().Trim());
                        entity.IsBureauVeritasPerformedAudit = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "IsBureauVeritasPerformedAudit")].Value).ToString().Trim());
                        entity.Dormitories = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "Dormitories")].Value).ToString().Trim());
                        entity.Canteen = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "Canteen")].Value).ToString().Trim());
                        entity.SecuirtyCameras = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "SecuirtyCameras")].Value).ToString().Trim());
                        entity.SecuirtyGuards = Convert.ToBoolean(GetBooleanValue(worksheet.Cells[row, GetColumnIndex(properties, "SecuirtyGuards")].Value).ToString().Trim());
                        list.Add(entity);
                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return list;
        }

        protected static int GetColumnIndex(string[] properties, string columnName)
        {
            if (properties == null)
            {
                throw new ArgumentNullException("properties");
            }

            if (columnName == null)
            {
                throw new ArgumentNullException("columnName");
            }

            for (int i = 0; i < properties.Length; i++)
            {
                if (properties[i].Equals(columnName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return i + 1; //// excel indexes start from 1
                }
            }

            return 0;
        }

        protected static object GetValue(object obj)
        {
            if (obj == null)
            {
                return string.Empty;
            }

            return obj;
        }

        protected static object GetDateValue(object obj)
        {
            if (obj == null)
            {
                return string.Empty;
            }

            return obj + "-01-01";
        }

        protected static object GetBooleanValue(object obj)
        {
            if (obj == null)
            {
                return false;
            }
            else
            {
                if (obj.ToString() == "True")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }
}
