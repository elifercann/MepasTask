using ClosedXML.Excel;
using DataAccess.Abstract;
using Entities.Concrete;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Repository
{
    public class CategoryRepository : ICategoryRepository
    {
        private string _pathToExcelFile;
        private string pathToExcelFile = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        public CategoryRepository(string pathToExcelFile)
        {
            _pathToExcelFile = pathToExcelFile;
        }
    
        public void AddCategory(Category category)
        {
           
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Categories"];
                worksheet.Equals(category);
                var end = worksheet.Dimension.End;

                var categoryId = end.Row + 1;

                worksheet.Cells[categoryId, 1].Value = category.id;
                worksheet.Cells[categoryId, 2].Value = category.name;


                package.Save();
            }
            }

        public void DeleteCategory(int id)
        {
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Categories"];

                // eşleşen idye sahip ürünü bulma
                var rowToDelete = worksheet.Cells["A:A"].FirstOrDefault(cell => cell.Value is int && (int)cell.Value ==id)?.Start.Row;
                if (rowToDelete == null)
                {
                    throw new ArgumentException("Kategori id'i bulunamadı", nameof(id));
                }

                // satırı sil
                worksheet.DeleteRow(rowToDelete.Value, 1);

                package.Save();
            }
        }

        public List<Category> GetCategories()
        {
            //lisans sorunundan dolayı aldığım hata da araştırma sonucu internette bulduğum çözüm kodu
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            //{
            //    var worksheet = package.Workbook.Worksheets["Categories"];
            //    var start = worksheet.Dimension.Start;
            //    var end = worksheet.Dimension.End;
            //    var categories = new List<Category>();

            //    for (var row = start.Row + 1; row <= end.Row; row++)
            //    {
            //        var category = new Category
            //        {
            //            id = worksheet.Cells[row, 1].GetValue<int>(),
            //            name = worksheet.Cells[row, 2].GetValue<string>(),

            //        };
            //        categories.Add(category);
            //    }

            //    return categories;
            //}
            var categories = new List<Category>();
            using (var workbook = new XLWorkbook(_pathToExcelFile))
            {
                var worksheet = workbook.Worksheet("Categories");

                // Find the column indexes for each property
                var idColumn = worksheet.Column("id");
                var nameColumn = worksheet.Column("adı");
               

                // Loop through each row in the worksheet (skip the header row)
                for (int row = 2; row <= worksheet.LastRow().RowNumber(); row++)
                {
                    var category = new Category
                    {
                        id = idColumn.Cell(row).GetValue<int>(),
                        name = nameColumn.Cell(row).GetValue<string>(),
                       
                    };
                    categories.Add(category);
                }
            }

            return categories;
        }

        public void UpdateCategory(Category category)
        {
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Products"];

                // Eşleşen idye sahip satırı bulma
                var rowToUpdate = worksheet.Cells["A:A"].FirstOrDefault(cell => cell.Value is int && (int)cell.Value == category.id)?.Start.Row;
                if (rowToUpdate == null)
                {
                    throw new ArgumentException("Kategori id'i bulunamadı", nameof(category));
                }

                // Satır için hücre değerlerini güncelleme
                worksheet.Cells[rowToUpdate.Value, 2].Value = category.name;
             

                package.Save();
            }
        }
    }
}
