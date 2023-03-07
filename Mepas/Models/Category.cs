namespace Mepas.Models
{
    public class Category
    {
        public string id { get; set; }
        public string name { get; set; }
        public virtual ICollection<Product> Products { get; set; }
    }
}
