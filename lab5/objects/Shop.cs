using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5.objects
{
    internal class Shop
    {
        public int ID { get; set; }
        public string Area { get; set; }
        public string Adress { get; set; }

        public Shop(int id, string area, string adress)
        {
            ID = id;
            Area = area;
            Adress = adress;
        }

        public override string ToString()
        {
            return $"ID: {ID}, район{Area}, адрес{Adress}";
        }
    }
}
