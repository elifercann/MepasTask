using ClosedXML.Excel;
using DataAccess.Repository;
using Entities.Concrete;
using Microsoft.AspNetCore.Mvc;

namespace Mepas.Controllers
{
    public class ProductController : Controller
    {
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veritabani.xlsx";
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
                        id = row.Cell(1).GetValue<int>(),
                        name = row.Cell(2).GetValue<string>(),

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
           
            return RedirectToAction(nameof(Index));
        }

    }
}
