using ExcelReaderUsingOpenOfficeXML.Models;
using RefreshOauth.Pluggins.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Script.Serialization;

namespace ExcelReaderUsingOpenOfficeXML.Controllers
{
    public class ExcelController : ApiController
    {
        ExcelServices excelServices = new ExcelServices();
        [Route("ExcelData")]
        [HttpPost]
        public IHttpActionResult ExcelData()
        {
           var usersList = new List<Student>();
            if (HttpContext.Current.Request.Files.AllKeys.Any())
            {
                // Get the uploaded image from the Files collection
                var file = HttpContext.Current.Request.Files["UploadedImage"];

                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    usersList = excelServices.ReadUploadedExcel(file);
                }

            }

            var json = new JavaScriptSerializer().Serialize(usersList);
            return Ok(usersList);
        }
    }
}
