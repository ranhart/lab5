using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab5.objects
{
    internal class Category
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string AgeRestriction { get; set; }

        public Category(int id, string name, string ageRestriction)
        {
            ID = id;
            Name = name;
            AgeRestriction = ageRestriction;
        }
        public Category(Category category)
        {
            ID = category.ID;
            Name = category.Name;
            AgeRestriction = category.AgeRestriction;
        }

        public override string ToString()
        {
            return $"ID: {ID}, название: {Name}, возрастное ограничение: {AgeRestriction}";
        }
    }
}
