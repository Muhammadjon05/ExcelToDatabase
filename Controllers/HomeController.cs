using System.Diagnostics;
using ExcelToDB.DbContext;
using Microsoft.AspNetCore.Mvc;
using ExcelToDB.Models;
using OfficeOpenXml;

namespace ExcelToDB.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly AppDbContext _context;
    public HomeController(ILogger<HomeController> logger, AppDbContext context)
    {
        _logger = logger;
        _context = context;
    }

    public IActionResult Index()
    {
        return View();
    }

    public async Task<IActionResult> ImportFromExcel(IFormFile file)
    {
        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);
            using (var package = new ExcelPackage(stream))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var rows = worksheet.Dimension.Rows;
                for (int i = 2; i < rows; i++)
                {
                     await _context.Students.AddAsync(new Student()
                    {
                        Name = worksheet.Cells[i, 1].Value.ToString()!.Trim(),
                        Surname = worksheet.Cells[i, 2].Value.ToString()!.Trim(),
                        ParentName = worksheet.Cells[i, 3].Value.ToString()!.Trim(),
                        Sex = worksheet.Cells[i, 4].Value.ToString()!.Trim(),
                        DateOfBirth = worksheet.Cells[i, 5].Value.ToString()!.Trim(),
                        Address = worksheet.Cells[i, 6].Value.ToString()!.Trim(),
                        Science = worksheet.Cells[i, 7].Value.ToString()!.Trim(),
                        Level = worksheet.Cells[i, 8].Value.ToString()!.Trim(),
                        Building = worksheet.Cells[i, 9].Value.ToString()!.Trim(),
                        SuitableDate = worksheet.Cells[i, 10].Value.ToString()!.Trim(),
                        SuitableTime = worksheet.Cells[i, 11].Value.ToString()!.Trim(),
                        ClassType = worksheet.Cells[i, 12].Value.ToString()!.Trim(),
                        PhoneNumber = worksheet.Cells[i, 13].Value.ToString()!.Trim(),
                    });
                     await _context.SaveChangesAsync();
                }
            }
        }

        return Ok("Done");
    }
    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}