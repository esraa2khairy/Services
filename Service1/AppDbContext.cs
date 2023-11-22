using Microsoft.EntityFrameworkCore;
using Service1.Models;
namespace Service1
{
    public class AppDbContext : DbContext
    {
        
            public AppDbContext(DbContextOptions<AppDbContext> options) : base(options)
            {
            }
           
            public DbSet<DataSheet> Data { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DataSheet>(e => { 
             e.Property(x => x.Id)
             .HasColumnName("Id")
             .ValueGeneratedNever();
        });
        }
    }
    
}
