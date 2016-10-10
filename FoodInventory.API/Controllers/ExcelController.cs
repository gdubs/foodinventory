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
using FoodInventory.API.Utilities;

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
            MemoryStream memoryStream = new MemoryStream(ExportService.ExcelReport(products));
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

                    ImportService<Product> imptService = new ImportService<Product>();
                    imptService.ValidateItems(file);

                    if (imptService._validRows.Count > 0)
                    {
                        foreach (var item in imptService._validRows)
                        {
                            _unitOfWork.ProductRepository.Insert(item);
                        }

                        _unitOfWork.Save();
                    }


                    return Request.CreateResponse(HttpStatusCode.OK, "Imported " + imptService._validRows.Count() + " rows with " + imptService._invalidRows.Count() + " invalid rows");
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
	}
}