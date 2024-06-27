using Microsoft.EntityFrameworkCore;
using System.Reflection;

namespace Ontap_Net104_319.Models
{
    public class AppDbContext: DbContext
    {
        public AppDbContext(){} // constructor ko tham số của nó
        public AppDbContext(DbContextOptions options) : base(options) // constructor kế thừa
        {
        }
        public DbSet<Account> Accounts { get; set; } // Đại diện cho 1 bảng trong CSDL
        public DbSet<Product> Products { get; set; }
        public DbSet<Bill> Bills { get; set; }
        public DbSet<BillDetails> BillDetails { get; set; }
        public DbSet<Cart> Carts { get; set; }
        public DbSet<CartDetails> CartDetails { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
			optionsBuilder.UseSqlServer("Data Source=ANHPHONGDZ;Initial Catalog=ontappp_PHONG;Integrated Security=True;TrustServerCertificate=true");
		}

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfigurationsFromAssembly(Assembly.GetExecutingAssembly());
        }
    }
}
