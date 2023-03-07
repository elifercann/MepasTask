namespace Mepas.Models
{
    public class Product
    {
        public int id { get; set; }
        public string name { get; set; }
        public int categoryId { get; set; }
        public decimal price { get; set; }
        public string unit { get; set; }
        public int stock { get; set; }
        public string color { get; set; }
        public decimal weight { get; set; }
        public decimal width { get; set; }
        public decimal height { get; set; }
        public int addedUserId { get; set; }
        public int updatedUserId { get; set; }
        public DateTime createdDate { get; set; }
        public DateTime updatedDate { get; set; }

        public Category categories { get; set; }
        public User addedUser { get; set; }
        public User updatedUser { get; set; }
    }
}
