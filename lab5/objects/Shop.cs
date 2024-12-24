using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5.objects
{
    internal class Shop
    {
        public string ID { get; set; }
        public string Area { get; set; }
        public string Adress { get; set; }

        public Shop(string id, string area, string adress)
        {
            ID = id;
            Area = area;
            Adress = adress;
        }

        public Shop(Shop shop)
        {
            ID = shop.ID;
            Area = shop.Area;
            Adress = shop.Adress;
        }

        public override string ToString()
        {
            return $"ID: {ID}, район{Area}, адрес{Adress}";
        }
    }
}
