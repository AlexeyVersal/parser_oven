using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace parser_oven.DB
{
    internal class Context : DbContext
    {
        public List<oc_product> products = new List<oc_product>();
        private string _connectionString = ConfigurationManager.AppSettings["ConnectionString"];
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseMySQL(_connectionString);
        }
        public DbSet<oc_product> oc_product { get; set; }
        public DbSet<oc_product_description> oc_product_description { get; set; }
        public async Task<List<oc_product>> GetListProductDB()
        {
            using (Context context = new Context())
            {
                products = context.oc_product.Where(p => p.manufacturer_id == 34).Include(p=>p.oc_Product_Description).ToList();
            }
            return products;
        }
    }
}
