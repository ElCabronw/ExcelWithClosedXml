using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelWithClosedXml.Models
{
    public class Sale
    {
        public int Id { get; set; }
        public string Product { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public decimal UnitValue{ get; set; }
        public decimal Total => Quantity * UnitValue;
        public DateTime SaleDate { get; set; }
        public string Seller { get; set; } = string.Empty;
    }
}
