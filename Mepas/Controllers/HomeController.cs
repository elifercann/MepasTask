using Mepas.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Drawing;

namespace Mepas.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        //public IActionResult CreateExcelProduct()
        //{
        //    var product = new List<Product>
        //    {
        //        new Product{id=1,name="Bulaşık Makinesi",categoryId=1 }
        //    };
        //    return View();
        //}
        //public IActionResult CreateExcelCategory()
        //{
        //    //categoryler newlenince oluturuldu
        //    var category = new List<Category>
        //    {
        //        new Category{id="1",name="Beyaz Eşya" },
        //        new Category{id="2",name="Küçük Ev Aletleri" },
        //        new Category{id="3",name="Teknolojik Aletler" },
        //    };
            
        //    var stream = new MemoryStream();
        //    //oluşturulan stream üzerinde çalışacak bu sayede kaydedip okuyabileceğiz
        //    using (var xlPackage =new ExcelPackage(stream))
        //    {
        //        var worksheet = xlPackage.Workbook.Worksheets.Add("Categories");
        //        worksheet.Cells["A1"].Value = "Kategoriler";
        //        //A1 ve B1 kolonlarını birleştirip görünümü ile ilgili işlemler sağlanıyor
        //        using (var t = worksheet.Cells["A1:B1"])
        //        {
        //            //birleştirilmesi için
        //            t.Merge = true;
        //            //font rengi
        //            t.Style.Font.Color.SetColor(Color.White);
        //            //yazının ortalanması için
        //            t.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
        //            //renk yerleşiminin nasıl olacağını ifade ediyor 
        //            t.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //            //arka plan rengi veriliyor
        //            t.Style.Fill.BackgroundColor.SetColor(Color.Gray);

        //        }
        //        //kategori id ve kategori isimleri için atama yapıldı
        //        worksheet.Cells["A3"].Value = "Id";
        //        worksheet.Cells["B3"].Value = "Adı";
        //        //düz bir renk vermek için
        //        worksheet.Cells["A3:B3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        //arka plan rengini veriliyor
        //        worksheet.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(Color.Azure);
        //        //başlıklar kalın font yapmak için
        //        worksheet.Cells["A3:B3"].Style.Font.Bold = true;

        //        //kategori bilgilerini girmek için 4.satırdan başlayacak
        //        int row = 4;
        //        //değerler veriliyor
        //        foreach (var item in category)
        //        {
        //            worksheet.Cells[row, 1].Value = item.id;
        //            worksheet.Cells[row, 2].Value = item.name;
        //            row++;
        //        }
        //        //stream üzerine kayıt yapılıyor
        //        xlPackage.Save();
        //        xlPackage.SaveAs(new FileInfo("wwwroot/Veritabani.xlsx"));
        //        return RedirectToAction("ReadCategory");
        //    }
        //}

        //public IActionResult ReadCategory()
        //{
        //    //dosya yolu veriliyor
        //    string path = "wwwroot/Veritabani.xlsx";
        //    //dosyayı fileinfo okuyacak
        //    FileInfo file=new FileInfo(path);

        //    ExcelPackage package = new ExcelPackage(file);
        //    //hangi worksheetini okuyacağını belirliyoruz
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets["Categories"];
        //    //ne kadar satır olacağını buluyoruz
        //    int rows = worksheet.Dimension.Rows;
        //    //ne kadar sütun olacağını buluyoruz
        //    int columns = worksheet.Dimension.Columns;  

        //    var categories=new List<Category>();
        //    //kategoriler 4.satırdan başladığı için rowu 4ten başlattık
        //    for (int i = 4; i <= rows; i++)
        //    {
        //        var item = new Category();
        //        item.id = worksheet.Cells[i, 1].Value.ToString();
        //        item.name = worksheet.Cells[i, 2].Value.ToString();

        //        categories.Add(item);
        //    }
        //    return View(categories);    
        //}
    }
}