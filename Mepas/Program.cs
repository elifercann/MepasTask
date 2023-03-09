using DataAccess.Abstract;
using DataAccess.Repository;
using Mepas.Controllers;
using OfficeOpenXml;
using System.Configuration;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddScoped<ICategoryRepository, CategoryRepository>();
builder.Services.AddScoped<IProductRepository, ProductRepository>();
builder.Services.AddHttpContextAccessor();
builder.Services.AddSession(options =>
{
    // Oturum süresi 30 dakika olarak ayarlanýr
    //options.IdleTimeout = TimeSpan.FromMinutes(30);

    // Cookie tabanlý oturum yönetimi
    options.Cookie.Name = "MyScheme";
    options.Cookie.MaxAge = TimeSpan.FromMinutes(60);
    options.Cookie.IsEssential = true;
});
builder.Services.AddSingleton("connectionString");
builder.Services.AddAuthentication("MyScheme")
    .AddCookie("MyScheme", options =>
    {
        options.LoginPath = "/Login/Login";
    });


var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();
app.UseSession(); // Oturum yönetimini etkinleþtirir
app.UseAuthentication(); // Kimlik doðrulamayý ekler
app.UseAuthorization();
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Login}/{action=Index}/{id?}");

app.Run();
