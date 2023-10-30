using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Technological_card
{
    internal class Component //2. Требования к материалам и комплектующим
    {
        static int count;
        [Name("Num")]
        int num;
        [Name("Name")]
        string name;
        [Name("Type")]
        string type;
        [Name("Unit")]
        string unit;
        [Name("Amount")]
        int amount; // using uint here

        public Component()
        {
            count++;
        }

        public Component (int num, string name, string type, string unit, int amount)
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.unit = unit;
            this.amount = amount;
            count++;
        }

        [Name("Num")]//[Name("№")]
        public int Num
        { get { return num; } set { num = value; } }
        [Name("Наименование")]
        public string Name
        { get { return name; } set { name = value; } }
        [Name("Тип (исполнение)")]
        public string Type
        { get { return type; } set { type = value; } }
        [Name("Ед. Изм.")]
        public string Unit
        { get { return unit; } set { unit = value; } }
        [Name("Кол-во")]
        public int Amount
        { get { return amount; } 
            set 
            {
                if (value > 0 ) { amount = value; } 
                else { Console.WriteLine("Кол-во не может быть 0 или отрицательным"); }

            } 
        }

        public string ToString() 
        {
            return $"{num} {name} {type} {unit} {amount}";
        }

    }
}
