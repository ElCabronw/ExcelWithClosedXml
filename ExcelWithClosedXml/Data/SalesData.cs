using ExcelWithClosedXml.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelWithClosedXml.Data
{
    public static class SalesData
    {
        public static List<Sale> GenerateFictitiousSales(int quantidade = 50)
        {
            var products = new[]
            {
                ("Dell Laptop", "Electronics", 3500.00m),
                ("Logitech Mouse", "Peripherals", 89.90m),
                ("Mechanical Keyboard", "Peripherals", 450.00m),
                ("LG 27\" Monitor", "Electronics", 1200.00m),
                ("Full HD Webcam", "Peripherals", 280.00m),
                ("Gaming Headset", "Peripherals", 320.00m),
                ("1TB SSD", "Hardware", 550.00m),
                ("16GB RAM", "Hardware", 380.00m),
                ("2TB External HDD", "Hardware", 420.00m),
                ("Gaming Chair", "Furniture", 1150.00m)
            };

            var vendedores = new[] { "João Silva", "Maria Santos", "Carlos Oliveira", "Ana Costa" };
            var random = new Random(42);

            var vendas = new List<Sale>();

            for (int i = 1; i <= quantidade; i++)
            {
                var produto = products[random.Next(products.Length)];

                vendas.Add(new Sale
                {
                    Id = i,
                    Product = produto.Item1,
                    Category = produto.Item2,
                    Quantity = random.Next(1, 20),
                    UnitValue = produto.Item3,
                    SaleDate = DateTime.Now.AddDays(-random.Next(0, 90)),
                    Seller = vendedores[random.Next(vendedores.Length)]
                });
            }

            return vendas.OrderByDescending(v => v.SaleDate).ToList();
        }
    }
}
