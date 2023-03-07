using ClosedXML.Excel;
using DataAccess.Repository;
using Entities.Concrete;
using Microsoft.AspNetCore.Mvc;

namespace Mepas.Controllers
{
    public class CategoryController : Controller
    {
       
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly string _sheetName = "Categories";

        public IActionResult Index()
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                var categories = worksheet.RowsUsed()
                    .Skip(1) // başlık satırını atlamak için
                    .Select(row => new Category
                    {
                        id = row.Cell(1).GetValue<int>(),
                        name = row.Cell(2).GetValue<string>(),

                    })
                    .ToList();
                return View(categories);
            }
        }

        public IActionResult Create()
        {
            
            return View();
        }

        [HttpPost]
        public IActionResult Create(Category category)
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
                        newRow.Cell(2).SetValue(category.name);

                        workbook.Save();
                    }
                    else
                    {
                        //boş ise gerekli hata yazılıyor
                        ModelState.AddModelError(string.Empty, $"Worksheet '{_sheetName}'çalışma kitabında bulunamadı.");
                        return View(category);
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {

                ModelState.AddModelError(string.Empty, "Kategori kaydedilirken hata oluştu: " + ex.Message);
                return View(category);
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
                var category = new Category
                {
                    id = row.Cell(1).GetValue<int>(),
                    name = row.Cell(2).GetValue<string>(),

                };
                return View(category);
            }
        }

        [HttpPost]
        public IActionResult Edit(int id, Category category)
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                var row = worksheet.RowsUsed().Skip(1)
                    .FirstOrDefault(r => r.Cell(1).GetValue<int>() == id);
                if (row == null)
                {
                    return NotFound();
                }
                row.Cell(2).SetValue(category.name);

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
