using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.Concrete
{
    public class Category
    {
        public int id { get; set; }
        public string name { get; set; }
        public virtual ICollection<Product> Products { get; set; }
    }
}
