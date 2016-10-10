using FoodInventory.Data.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;

namespace FoodInventory.API.Utilities
{
    public class ExportService
    {
        static public byte[] ExcelReport(List<ProductDTO> products)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Report");
                // load headers

                var columns = new List<string>();
                foreach (PropertyInfo prop in typeof(ProductDTO).GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if (prop.Name != "DeletedDate")
                        columns.Add(prop.Name);
                }
                for (var c = 0; c < columns.Count; c++)
                {
                    sheet.Cells[1, c + 1].Value = columns[c];
                }

                for (var p = 0; p < products.Count; p++)
                {
                    var product = products[p];
                    for (var c = 0; c < columns.Count; c++)
                    {
                        sheet.Cells[p + 2, c + 1].Value = product.GetType().GetProperty(columns[c]).GetValue(product, null);
                    }
                }

                return package.GetAsByteArray();
            }
        }
    }
}