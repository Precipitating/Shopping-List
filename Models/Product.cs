using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.Globalization;

namespace ShoppingList.Models
{
    public class Product
    {
        public int Id { get; set; }
        [MaxLength(200)]
        public string Name { get; set; } = "";
        [MaxLength(100)]
        public string Brand { get; set; } = "";
        [MaxLength(100)]
        public string Category { get; set; } = "";
        [Precision(16,2)]
        public decimal Price { get; set; }
        [MaxLength(100)]
        public string Description { get; set; } = "";
        [MaxLength(100)]
        public string ImageFileName { get; set; } = "";

        public DateTime Created { get; set; }


        public static string GetCurrencySymbol()
        {
            CultureInfo culture = CultureInfo.CurrentCulture;
            return culture.NumberFormat.CurrencySymbol;
        }




    }
}
