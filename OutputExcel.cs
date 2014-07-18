namespace VAU.Web.CommonCode
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Reflection;

    /// <summary>
    /// Summary description for OutputExcel
    /// </summary>
    public class OutputExcel
    {
        #region Public Methods and Operators

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
        #endregion
    }
}
