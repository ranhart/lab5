using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5
{
    internal class Product
    {
        public int Article { get; set; }
        public int CategoryID { get; set; }
        public string ItemName { get; set; }
        public int FirstPrice { get; set; }
        public int SecondPrice { get; set; }
        public string Discount { get; set; }

        public Product(int article, int categoryID, string itemName, int firstPrice, int secondPrice, string discount)
        {
            Article = article;
            CategoryID = categoryID;
            ItemName = itemName;
            FirstPrice = firstPrice;
            SecondPrice = secondPrice;
            Discount = discount;
        }

        public override string ToString()
        {
            return $"Артикул: {Article}, ID категории: {CategoryID}, название товара: {ItemName}, цена закупки при поступлении: {FirstPrice}, цена продажи без скидки: {SecondPrice}, скидка: {Discount}";
        }
    }
}
