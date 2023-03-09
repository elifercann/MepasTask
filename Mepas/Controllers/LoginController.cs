using ClosedXML.Excel;
using Entities.Concrete;
using Mepas.Models;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Mvc;
using User = Entities.Concrete.User;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;



namespace Mepas.Controllers
{
    public class LoginController : Controller
    {
        private readonly string _excelPath = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        private readonly int _usernameColumnIndex = 4; // Excel dosyasındaki kullanıcı adı sütununun index'i
        private readonly int _passwordColumnIndex = 5; // Excel dosyasındaki şifre sütununun index'i
        private readonly string _sheetName = "Users";
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly ILogger<LoginController> _logger;

        public LoginController(ILogger<LoginController> logger, IHttpContextAccessor httpContextAccessor)
        {
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
        }
        [Authorize]
        public IActionResult Index()
        {
            return View();
        }


        [HttpGet]
        public ActionResult Login()
        {
            return View();
        }
        //çalışan kodlar
        [HttpPost]
        public ActionResult Login(string username, string password)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // EPPlus kullanarak Excel dosyasından kullanıcı adı ve şifreleri kontrol edin

            using (var package = new ExcelPackage(new FileInfo(_excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[_sheetName];
                var rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) // Excel dosyasında ilk satır başlık olduğu için 2. satırdan başlayın
                {
                    var excelUsername = worksheet.Cells[row, _usernameColumnIndex].Value?.ToString();
                    var excelPassword = worksheet.Cells[row, _passwordColumnIndex].Value?.ToString();

                    if (username == excelUsername && password == excelPassword)
                    {
                        // Kullanıcı adı ve şifre doğruysa session'a kaydedin ve Home sayfasına yönlendirin
                        _httpContextAccessor.HttpContext.Session.SetString("username", username);
                       
                        return RedirectToAction("Index", "Product");
                    }
                }

                // Kullanıcı adı ve şifre yanlışsa hata mesajı gösterin
                ViewBag.ErrorMessage = "Kullanıcı adı veya şifre yanlış.";
                return View();
            }
        }

        public ActionResult Logout()
        {
            // Session'ı sıfırlayın ve Login sayfasına yönlendirin
            _httpContextAccessor.HttpContext.Session.Clear();
            return RedirectToAction("Login", "Login");
        }

    }


}

