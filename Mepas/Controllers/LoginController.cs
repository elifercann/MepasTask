using ClosedXML.Excel;
using Entities.Concrete;
using Mepas.Models;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using User = Entities.Concrete.User;

namespace Mepas.Controllers
{
    public class LoginController : Controller
    {
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly string _sheetName = "Users";
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult Index(LoginViewModel model)
        {
            if (ModelState.IsValid)
            {
                var user = GetUserByUsername(model.username);

                if (user != null && user.password == model.password)
                {
                    // Authentication successful, redirect to home page
                    return RedirectToAction("Index", "Home");
                }
                else if (string.IsNullOrEmpty(model.username))
                {
                    ModelState.AddModelError("Username", "Username is required.");
                }
                else
                {
                    ModelState.AddModelError("Password", "Incorrect password.");
                }
            }

            return RedirectToAction("Index","Product");
        }
        public IActionResult Logout()
        {
            HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
            return RedirectToAction("Index");
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
