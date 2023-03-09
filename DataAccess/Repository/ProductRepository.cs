using DataAccess.Abstract;
using Entities.Concrete;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Repository
{
    public class ProductRepository:IProductRepository
    {
        private string _pathToExcelFile;
        private string pathToExcelFile = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        public ProductRepository(string pathToExcelFile)
        {
            _pathToExcelFile = pathToExcelFile;
        }

     
        public List<Product> GetAllProducts()
        {
            //lisans sorunundan dolayı aldığım hata da araştırma sonucu internette bulduğum çözüm kodu
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Products"];
                //null exception hatası
                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;
                var products = new List<Product>();

                for (var row = start.Row + 1; row <= end.Row; row++)
                {
                    var product = new Product
                    {
                        id = worksheet.Cells[row, 1].GetValue<int>(),
                        name = worksheet.Cells[row, 2].GetValue<string>(),
                        categoryId = worksheet.Cells[row, 3].GetValue<int>(),
                        price = worksheet.Cells[row, 4].GetValue<int>(),
                        unit = worksheet.Cells[row, 5].GetValue<string>(),
                        stock = worksheet.Cells[row, 6].GetValue<int>(),
                        color = worksheet.Cells[row, 7].GetValue<string>(),
                        weight = worksheet.Cells[row, 8].GetValue<decimal>(),
                        width = worksheet.Cells[row, 9].GetValue<decimal>(),
                        height = worksheet.Cells[row, 10].GetValue<decimal>(),
                        addedUserId = worksheet.Cells[row, 11].GetValue<int>(),
                        updatedUserId = worksheet.Cells[row, 12].GetValue<int>(),
                        createdDate = worksheet.Cells[row, 13].GetValue<DateTime>(),
                        updatedDate = worksheet.Cells[row, 14].GetValue<DateTime>()
                    };
                    products.Add(product);
                }

                return products;
            }
        }
      

        public void AddProduct(Product product)
        {
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Products"];
                var end = worksheet.Dimension.End;

                var newProductId = end.Row + 1;

                worksheet.Cells[newProductId, 1].Value = product.id;
                worksheet.Cells[newProductId, 2].Value = product.name;
                worksheet.Cells[newProductId, 3].Value = product.categoryId;
                worksheet.Cells[newProductId, 4].Value = product.price;
                worksheet.Cells[newProductId, 5].Value = product.unit;
                worksheet.Cells[newProductId, 6].Value = product.stock;
                worksheet.Cells[newProductId, 7].Value = product.color;
                worksheet.Cells[newProductId, 8].Value = product.weight;
                worksheet.Cells[newProductId, 9].Value = product.width;
                worksheet.Cells[newProductId, 10].Value = product.height;
                worksheet.Cells[newProductId, 11].Value = product.addedUserId;
                worksheet.Cells[newProductId, 12].Value = product.updatedUserId;
                worksheet.Cells[newProductId, 13].Value = product.createdDate;
                worksheet.Cells[newProductId, 14].Value = product.updatedDate;

                package.Save();
            }
        }

        public void DeleteProduct(int productId)
        {
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Products"];

                // eşleşen idye sahip ürünü bulma
                var rowToDelete = worksheet.Cells["A:A"].FirstOrDefault(cell => cell.Value is int && (int)cell.Value == productId)?.Start.Row;
                if (rowToDelete == null)
                {
                    throw new ArgumentException("Product id'i bulunamadı", nameof(productId));
                }

                // satırı sil
                worksheet.DeleteRow(rowToDelete.Value, 1);

                package.Save();
            }
        }

        public void UpdateProduct(Product product)
        {
            using (var package = new ExcelPackage(new FileInfo(_pathToExcelFile)))
            {
                var worksheet = package.Workbook.Worksheets["Products"];

                // Eşleşen idye sahip satırı bulma
                var rowToUpdate = worksheet.Cells["A:A"].FirstOrDefault(cell => cell.Value is int && (int)cell.Value == product.id)?.Start.Row;
                if (rowToUpdate == null)
                {
                    throw new ArgumentException("Product id'i bulunamadı", nameof(product));
                }

                // Satır için hücre değerlerini güncelleme
                worksheet.Cells[rowToUpdate.Value, 2].Value = product.name;
                worksheet.Cells[rowToUpdate.Value, 3].Value = product.categoryId;
                worksheet.Cells[rowToUpdate.Value, 4].Value = product.price;
                worksheet.Cells[rowToUpdate.Value, 5].Value = product.unit;
                worksheet.Cells[rowToUpdate.Value, 6].Value = product.stock;
                worksheet.Cells[rowToUpdate.Value, 7].Value = product.color;
                worksheet.Cells[rowToUpdate.Value, 8].Value = product.weight;
                worksheet.Cells[rowToUpdate.Value, 9].Value = product.width;
                worksheet.Cells[rowToUpdate.Value, 10].Value = product.height;
                worksheet.Cells[rowToUpdate.Value, 11].Value = product.addedUserId;
                worksheet.Cells[rowToUpdate.Value, 12].Value = product.updatedUserId;
                worksheet.Cells[rowToUpdate.Value, 13].Value = product.createdDate;
                worksheet.Cells[rowToUpdate.Value, 14].Value = product.updatedDate;

                package.Save();
            }
        }

      
    }
}
