using ClosedXML.Excel;
using DataAccess.Abstract;
using DataAccess.Repository;
using DocumentFormat.OpenXml.Office2010.Excel;
using Entities.Concrete;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Mepas.Controllers
{
   
    public class CategoryController : Controller
    {

        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly string _sheetName = "Categories";
        //private readonly ICategoryRepository _categoryRepository;
        //private readonly IProductRepository _productRepository;

        //public CategoryController(ICategoryRepository categoryRepository, IProductRepository productRepository)
        //{
        //    _categoryRepository = categoryRepository;
        //    _productRepository = productRepository;
        //}

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
                        //newRow.Cell(2).SetValue(category.Products);

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
        //hatalı alanlar bakılacak
        //public IActionResult GetByProductList(int categoryId)
        //{
        //    using (var workbook = new XLWorkbook(_fileName))
        //    {
        //        var worksheet = workbook.Worksheet(_sheetName);
        //        var categories = worksheet.RowsUsed()
        //                 .Skip(1) // başlık satırını atlamak için
        //                    .Select(row => new Category
        //                   {
        //                      id = row.Cell(1).GetValue<int>(),
        //                      name = row.Cell(2).GetValue<string>(),
        //                      //categoryId = int.TryParse(row.Cell(3).Value.ToString(), out int result) ? result : 0,
        //                    })
        //                     .ToList();
        //        var productsByCategory = categories.GroupBy(p => p.id);
        //        return View(productsByCategory);

        //    }


        //}
        //hatalı kod
        //public IActionResult ExportProductsToExcel()
        //{
        //    using (var workbook = new XLWorkbook())
        //    {
        //        var worksheet = workbook.Worksheets.Add("Products");

        //        // Products listesi alınır
        //        var products = _productRepository.GetAllProducts();

        //        // Tabloya başlık satırı eklenir
        //        var headerRow = new List<string>() { "ID", "Name", "Category", "Price" };
        //        worksheet.Cell(1, 1).InsertTable(new List<List<string>> { headerRow });

        //        // Tabloya veriler eklenir
        //        var dataRows = products.Select(p => new List<string>()
        //    {
        //        p.id.ToString(),
        //        p.name,
        //        p.categories.name,
        //        p.price.ToString()
        //    }).ToList();

        //        worksheet.Cell(2, 1).InsertTable(dataRows);

        //        // Dosya kaydedilir
        //        var stream = new MemoryStream();
        //        workbook.SaveAs(stream);
        //        stream.Seek(0, SeekOrigin.Begin);
        //        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Products.xlsx");
        //    }
        //}
    }
}
