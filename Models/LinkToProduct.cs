﻿using System.ComponentModel.DataAnnotations;

namespace ShoppingList.Models
{
    public enum ACTION_TYPE
    {
        ALL,
        PRICE
    }
    public class LinkToProduct
    {
        [Required] public string Link { get; set; } = "";

        [Required, MaxLength(100)]
        public string Category { get; set; } = "";
        [Required, MaxLength(100)]
        public string Description { get; set; } = "";


    }
}
