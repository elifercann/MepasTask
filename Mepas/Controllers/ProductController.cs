using ClosedXML.Excel;
using DataAccess.Repository;
using Entities.Concrete;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Mepas.Controllers
{
 
    public class ProductController : Controller
    {
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly string _sheetName = "Products";



        public IActionResult Index()
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                var products = worksheet.RowsUsed()
                    .Skip(1) // Skip the header row
                    .Select(row => new Product
                    {
                        //propertyler hangi alandaysa tek tek alınıyor tipin göre listeleniyor
                        id = row.Cell(1).GetValue<int>(),
                        name = row.Cell(2).GetValue<string>(),
                        categoryId = row.Cell(3).GetValue<int>(),
                        price = row.Cell(4).GetValue<decimal>(),
                        unit = row.Cell(5).GetValue<string>(),
                        stock = row.Cell(6).GetValue<int>(),
                        color = row.Cell(7).GetValue<string>(),
                        weight = row.Cell(8).GetValue<decimal>(),
                        width = row.Cell(9).GetValue<decimal>(),
                        height = row.Cell(10).GetValue<decimal>(),
                       
                    })
                    .ToList();
                return View(products);
            }
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Create(Product product)
        {

            try
            {
                using (var workbook = new XLWorkbook(_fileName))
                {
                    //sheet ismine göre değerleri eklenecek alanlar belirleniyor
                    var worksheet = workbook.Worksheet(_sheetName);
                    //sheet boş değilse ekleme işlemi yapılıyor
                    if (worksheet != null)
                    {
                        //en son kullanılan satırı bulmak için 
                        var lastRow = worksheet.LastRowUsed().RowNumber();
                        //eklenecek alan için son satırdan sonrası ayarlanıyor
                        var newRow = worksheet.Row(lastRow + 1);
                        newRow.Cell(1).SetValue(lastRow);
                        newRow.Cell(2).SetValue(product.name);
                        newRow.Cell(3).SetValue(product.categoryId);
                        newRow.Cell(4).SetValue(product.price);
                        newRow.Cell(5).SetValue(product.unit);
                        newRow.Cell(6).SetValue(product.stock);
                        newRow.Cell(7).SetValue(product.color);
                        newRow.Cell(8).SetValue(product.weight);
                        newRow.Cell(9).SetValue(product.width);
                        newRow.Cell(10).SetValue(product.height);
                        newRow.Cell(11).SetValue(product.addedUserId);
                        newRow.Cell(12).SetValue(product.updatedUserId);
                        newRow.Cell(13).SetValue(product.createdDate);
                        newRow.Cell(14).SetValue(product.updatedDate);

                        workbook.Save();
                    }
                    else
                    {
                        //boş ise gerekli hata yazılıyor
                        ModelState.AddModelError(string.Empty, $"Worksheet '{_sheetName}'çalışma kitabında bulunamadı.");
                        return View(product);
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {

                ModelState.AddModelError(string.Empty, "Ürün kaydedilirken hata oluştu: " + ex.Message);
                return View(product);
            }

        }
        public IActionResult Edit(int id)
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                //seçilen hücre için idye göre verilerin getirilmesi için gerekli kod
                var row = worksheet.RowsUsed().Skip(1)
                    .FirstOrDefault(r => r.Cell(1).GetValue<int>() == id);
                if (row == null)
                {
                    return NotFound();
                }
                var product = new Product
                {
                    id = row.Cell(1).GetValue<int>(),
                    name = row.Cell(2).GetValue<string>(),
                    categoryId = row.Cell(3).GetValue<int>(),
                    price = row.Cell(4).GetValue<decimal>(),
                    unit = row.Cell(5).GetValue<string>(),
                    stock = row.Cell(6).GetValue<int>(),
                    color = row.Cell(7).GetValue<string>(),
                    weight = row.Cell(8).GetValue<decimal>(),
                    width = row.Cell(9).GetValue<decimal>(),
                    height = row.Cell(10).GetValue<decimal>(),
                    //addedUserId = row.Cell(11).GetValue<int>(),
                    updatedUserId = row.Cell(12).GetValue<int>(),
                    //createdDate = row.Cell(13).GetValue<DateTime>(),
                    updatedDate = row.Cell(14).GetValue<DateTime>(),
                   

                };
                return View(product);
            }
        }

        [HttpPost]
        public IActionResult Edit(int id, Product product)
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                var row = worksheet.RowsUsed().Skip(1)//ilk satırı atlıyor,ilk satır başlık
                    .FirstOrDefault(r => r.Cell(1).GetValue<int>() == id);
                if (row == null)
                {
                    return NotFound();
                }
                row.Cell(2).SetValue(product.name);
                row.Cell(3).SetValue(product.categoryId);
                row.Cell(4).SetValue(product.price);
                row.Cell(5).SetValue(product.unit);
                row.Cell(6).SetValue(product.stock);
                row.Cell(7).SetValue(product.color);
                row.Cell(8).SetValue(product.weight);
                row.Cell(9).SetValue(product.width);
                row.Cell(10).SetValue(product.height);
                //row.Cell(11).SetValue(product.addedUserId);
                row.Cell(12).SetValue(product.updatedUserId);
                //row.Cell(13).SetValue(product.createdDate);
                row.Cell(14).SetValue(product.updatedDate);

                workbook.Save();
            }
            return RedirectToAction(nameof(Index));
        }


        public IActionResult Delete(int id)
        {
            try
            {
                using (var workbook = new XLWorkbook(_fileName))
                {
                    var worksheet = workbook.Worksheet(_sheetName);

                    // silinecek verinin idye göre getirilmesi
                    var rowToDelete = worksheet.RowsUsed().Skip(1)
                        .FirstOrDefault(row => row.Cell(1).GetValue<int>() == id);

                    if (rowToDelete != null)
                    {
                        rowToDelete.Delete();
                        workbook.Save();
                    }
                    else
                    {
                        ModelState.AddModelError(string.Empty, $"idye ait {id} çalışma kitabında bulunamadı..");
                    }
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError(string.Empty, "Kayıt silinirken hata oluştu " + ex.Message);
            }


            return RedirectToAction(nameof(Index));
        }
    }

}

