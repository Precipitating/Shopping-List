using Microsoft.EntityFrameworkCore;
using ShoppingList.Models;

namespace ShoppingList.Services
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions options) : base(options)
        {
        
        
        }

        public DbSet<Product> Products { get; set; }
    }
}
