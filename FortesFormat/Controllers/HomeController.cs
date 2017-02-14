using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Excel;
using System.Data;
using System.Text;
using System;

namespace FortesFormat.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            return View();
        }

        [HttpPost]
        public ActionResult convertExcel(FormCollection form)
        {
            try
            {
                string arqAux = Server.MapPath("~/midia/arq.csv");
                HttpPostedFileBase file = Request.Files["UploadedFile"];

                //verifico se o arquivo veio no request
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    //aqui eu salvo o arquivo excel no servidor para poder manipulá-lo
                    string fileName = Path.GetFileName(file.FileName);
                    string location = Path.Combine(Server.MapPath("~/midia"), fileName);
                    file.SaveAs(location);

                    FileStream stream = System.IO.File.Open(Server.MapPath("~/midia/" + fileName + ""), FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    //transforma o excel em dataset
                    DataSet excel = excelReader.AsDataSet();

                    StringBuilder sb = new StringBuilder();

                    //adiciona as linhas
                    foreach (DataRow row in excel.Tables[0].Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(";", fields));
                    }

                    //aqui eu passo o stringBiulder para o meu arquivo auxiliar
                    System.IO.File.WriteAllText(arqAux, sb.ToString(), Encoding.UTF8);

                    //deleto o arquivo pdf que tinha salvado
                    System.IO.File.Delete(Server.MapPath("~/midia/" + fileName + ""));

                    return File(arqAux, "application/csv", "" + Path.GetFileNameWithoutExtension(fileName) + ".csv");
                }
                else
                {
                    return RedirectToAction("index");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return View("Error");
            }

        }
    }
}