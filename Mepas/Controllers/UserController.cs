using ClosedXML.Excel;
using Entities.Concrete;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Mepas.Controllers
{
   
    public class UserController : Controller
    {
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly string _sheetName = "Users";
        public IActionResult Index()
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                var users = worksheet.RowsUsed()
                    .Skip(1) // başlık satırını atlamak için
                    .Select(row => new User
                    {
                        id = row.Cell(1).GetValue<int>(),
                        name = row.Cell(2).GetValue<string>(),
                        surname = row.Cell(3).GetValue<string>(),
                        username = row.Cell(4).GetValue<string>(),
                        password = row.Cell(5).GetValue<string>(),
                        status = row.Cell(6).GetValue<bool>(),
                        

                    })
                    .ToList();
                return View(users);
            }
        }
        public IActionResult Create()
        {

            return View();
        }

        [HttpPost]
        public IActionResult Create(User user)
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
                        newRow.Cell(2).SetValue(user.name);
                        newRow.Cell(3).SetValue(user.surname);
                        newRow.Cell(4).SetValue(user.username);
                        newRow.Cell(5).SetValue(user.password);
                        newRow.Cell(6).SetValue(user.status);

                        workbook.Save();
                    }
                    else
                    {
                        //boş ise gerekli hata yazılıyor
                        ModelState.AddModelError(string.Empty, $"Worksheet '{_sheetName}'çalışma kitabında bulunamadı.");
                        return View(user);
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {

                ModelState.AddModelError(string.Empty, "Kullanıcı kaydedilirken hata oluştu: " + ex.Message);
                return View(user);
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
                var user = new User
                {
                    id = row.Cell(1).GetValue<int>(),
                    name = row.Cell(2).GetValue<string>(),
                    surname = row.Cell(3).GetValue<string>(),
                    username = row.Cell(4).GetValue<string>(),
                    password = row.Cell(5).GetValue<string>(),
                    status = row.Cell(6).GetValue<bool>(),

                };
                return View(user);
            }
        }

        [HttpPost]
        public IActionResult Edit(int id, User user)
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
                row.Cell(2).SetValue(user.name);
                row.Cell(3).SetValue(user.surname);
                row.Cell(4).SetValue(user.username);
                row.Cell(5).SetValue(user.password);
                row.Cell(6).SetValue(user.status);

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

        public User GetUserByUsername(string username)
        {
            using (var workbook = new XLWorkbook(_fileName))
            {
                var worksheet = workbook.Worksheet(_sheetName);
                if (worksheet.IsEmpty())
                {
                    return null;
                }

                var row = worksheet.RowsUsed()
                    .Skip(1) // başlık atlanıyor
                    .FirstOrDefault(r => r.Cell(3).Value.ToString() == username);//username 4.hücrede

                if (row == null)
                {
                    return null;
                }

                return new User
                {
                    name = row.Cell(1).Value.ToString(),
                    surname = row.Cell(2).Value.ToString(),
                    username = row.Cell(3).Value.ToString(),
                    password = row.Cell(4).Value.ToString(),
                    status = bool.Parse(row.Cell(5).Value.ToString())
                };
            }
        }
    }
}
