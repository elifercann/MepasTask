using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entities.Concrete
{
    public class User
    {
        public int id { get; set; }
        public string name { get; set; }
        public string surname { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public bool status { get; set; }

        public virtual ICollection<Product> AddedProducts { get; set; }
        public virtual ICollection<Product> UpdatedProducts { get; set; }
    }
}
