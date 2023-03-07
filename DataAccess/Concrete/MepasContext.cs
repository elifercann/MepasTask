using Entities.Concrete;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Concrete
{
    public class MepasContext:DbContext
    {
        public MepasContext(DbContextOptions<MepasContext> options)
        : base(options)
        {
        }
       
        public DbSet<Product> Products { get; set; }
        public DbSet<Category> Categories { get; set; }
        public DbSet<User> Users { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            base.OnConfiguring(optionsBuilder);

            //string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\to\your\file.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES;'";
            //optionsBuilder.UseOleDb(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Product>()
                .HasOne(p => p.categories)
                .WithMany(c => c.Products)
                .HasForeignKey(p => p.categoryId);

            modelBuilder.Entity<Product>()
                .HasOne(p => p.addedUser)
                .WithMany(u => u.AddedProducts)
                .HasForeignKey(p => p.addedUserId);

            modelBuilder.Entity<Product>()
                .HasOne(p => p.updatedUser)
                .WithMany(u => u.UpdatedProducts)
                .HasForeignKey(p => p.updatedUserId);
        }
    }
}

