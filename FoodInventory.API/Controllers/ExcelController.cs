using Excel;
using FoodInventory.API.Utilities;
using FoodInventory.Data;
using FoodInventory.Data.Models;
using FoodInventory.Data.Models.DTOs;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.Http;

namespace FoodInventory.API.Controllers
{
    [RoutePrefix("api/Product/Excel")]
    public class ExcelController : ApiController
    {

        private UnitOfWork _unitOfWork = new UnitOfWork();

        //
        // GET: /Excel/
        [Route("Export")]
        [HttpGet]
        public HttpResponseMessage Export()
        {

            var response = new HttpResponseMessage(HttpStatusCode.OK);

            List<ProductDTO> products = _unitOfWork.ProductRepository
                                                    .Get()
                                                    .Select(p => new ProductDTO()
                                                            {
                                                                ID = p.ID,
                                                                Name = p.Name,
                                                                Description = p.Description,
                                                                PurchasePrice = p.PurchasePrice,
                                                                SalesPrice = p.SalesPrice,
                                                                SpoilDate = p.SpoilDate,
                                                                UnitsAvailable = p.UnitsAvailable
                                                            }).ToList();

            MediaTypeHeaderValue mediaType = new MediaTypeHeaderValue("application/octet-stream");
            MemoryStream memoryStream = new MemoryStream(ExcelReport(products));
            response.Content = new StreamContent(memoryStream);
            response.Content.Headers.ContentType = mediaType;
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("fileName") { FileName = "myReport.xls" };
            
            return response;
        }
        [HttpPost]
        [Route("Import")]
        public HttpResponseMessage Import()
        {

            try
            {
                if (HttpContext.Current.Request.Files.Count > 0)
                {
                    var file = HttpContext.Current.Request.Files[0];
                    var fileName = Path.GetFileName(file.FileName);

                    Import<Product> xl = new Import<Product>(file);
                    //var items = xl.ValidateItems(file);

                    if (xl._validRows.Count > 0)
                    {
                        foreach (var item in xl._validRows)
                        {
                            _unitOfWork.ProductRepository.Insert(item);
                        }

                        _unitOfWork.Save();
                    }


                    return Request.CreateResponse(HttpStatusCode.OK, "Imported " + xl._validRows.Count() + " rows with " + xl._invalidRows.Count() + " invalid rows");
                }
                else
                {
                    return Request.CreateResponse(HttpStatusCode.BadRequest, "No file found.");
                }
            }
            catch (Exception ex)
            {
                return Request.CreateResponse(HttpStatusCode.BadRequest, ex.Message);
            } 
        }

        public byte[] ExcelReport(List<ProductDTO> products)
        {
            using (var package = new ExcelPackage()){
                var sheet = package.Workbook.Worksheets.Add("Report");
                // load headers

                var columns = new List<string>();
                foreach (PropertyInfo prop in typeof(ProductDTO).GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if(prop.Name != "DeletedDate")
                        columns.Add(prop.Name);
                }
                for (var c = 0; c < columns.Count; c++)
                {
                    sheet.Cells[1, c + 1].Value = columns[c];
                }

                for(var p = 0; p < products.Count; p++){
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