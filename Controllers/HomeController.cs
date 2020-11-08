using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace SpreadSheet.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult Baixar()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID Cliente");
            
            FileInfo fileInfoTemplate = new FileInfo(System.Web.HttpContext.Current.Server.MapPath("~/Template/Template.xlsx"));

            OfficeOpenXml.ExcelPackage excel = new ExcelPackage(fileInfoTemplate);

            ExcelWorksheet worksheet = excel.Workbook.Worksheets.Add("BloqueioClientes");
            worksheet.Cells["A1"].LoadFromDataTable(dt, true);

            MemoryStream stream = new MemoryStream();
            excel.SaveAs(stream);


            HttpContext.Response.Clear();
            HttpContext.Response.AddHeader("content-disposition", string.Format("attachment;filename=BloqueioClientes.xlsx"));

            HttpContext.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Response.ContentEncoding = System.Text.Encoding.Default;

            HttpContext.Response.Cache.SetCacheability(HttpCacheability.NoCache);

            stream.WriteTo(Response.OutputStream);

            Response.End();

            return null;
        }
    }
}