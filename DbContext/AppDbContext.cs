using ExcelToDB.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelToDB.DbContext;

public class AppDbContext : Microsoft.EntityFrameworkCore.DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
    {
        
    }
    public DbSet<Student> Students { get; set; }
}