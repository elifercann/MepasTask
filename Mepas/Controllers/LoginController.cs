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
using Microsoft.Office.Interop.Excel;
using Microsoft.AspNetCore.Authorization;

namespace Mepas.Controllers
{
    public class LoginController : Controller
    {
        private readonly string _fileName = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
        //string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\ercan\\Desktop\\task\\veri2.xlsx;Extended Properties='Excel 12.0;HDR=YES;'";

        private readonly string _sheetName = "Users";
        private readonly ILogger<LoginController> _logger;

        public LoginController(ILogger<LoginController> logger)
        {
            _logger = logger;
        }
         [Authorize]
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult Login()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Login(LoginViewModel model)
        {
            var filePath = @"C:\Users\ercan\Desktop\task\veri2.xlsx";
            var app = new Application();
            var workbook = app.Workbooks.Open(filePath);
            Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);
            var range = worksheet.UsedRange;
            

            for (int row = 2; row <= range.Rows.Count; row++)
            {
                var usernameCell = (Microsoft.Office.Interop.Excel.Range)range.Cells[row, 1];
                var passwordCell = (Microsoft.Office.Interop.Excel.Range)range.Cells[row, 2];
                var username = usernameCell.Value2?.ToString();
                var password = passwordCell.Value2?.ToString();

                if (usernameCell== null || passwordCell == null)
                {
                    continue;
                }
              

                if (model.username == username && model.password == password)
                {
                    var claims = new List<Claim>
                    {
                        new Claim(ClaimTypes.Name, model.username)
                    };
                    var identity = new ClaimsIdentity(claims, "login");
                    var principal = new ClaimsPrincipal(identity);
                    await HttpContext.SignInAsync(principal);
                    return RedirectToAction("Index", "Home");
                }
            }

            ModelState.AddModelError("", "Invalid login attempt.");
            return View();
        }

        public async Task<IActionResult> Logout()
        {
            await HttpContext.SignOutAsync();
            return RedirectToAction("Login", "Login");
        }

      
    }
}

