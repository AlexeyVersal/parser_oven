using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace parser_oven.DB
{
    internal class oc_product
    {
        [Key]
        public int product_id { get; set; }
        public string model { get; set; }
        public string sku { get; set; }
        public int manufacturer_id { get; set; }
        public double price { get; set; }
        public double benchmark_price { get; set; }
        [ForeignKey("product_id")]
        public oc_product_description oc_Product_Description { get; set; }

    }
    public class oc_product_description
    {
        [Key]
        [ForeignKey("product_id")]
        public int product_id { get; set; }
        public string name { get; set; }
    }
    public class ProductApi
    {
        public string client_id = ConfigurationManager.AppSettings["client_id"];
        public List<ProductFind> products { get; set; } = new List<ProductFind>();
    }
    public class ProductFind
    {
        public int product_id { get; set; }
        public string price { get; set; }
    }
}
