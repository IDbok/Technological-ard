
using System.Text.Json.Serialization;

namespace Technological_card
{
    internal class Component : Struct//2. Требования к материалам и комплектующим
    {
        static int count;
        int num;
        string name;
        string type;
        string unit;
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
        [JsonInclude]
        public override int Num
        { get { return num; } set { num = value; } }
        [JsonInclude]
        public override string Name
        { get { return name; } set { name = value; } }
        [JsonInclude]
        public override string Type
        { get { return type; } set { type = value; } }
        public override string Unit
        { get { return unit; } set { unit = value; } }
        public override int Amount
        { get { return amount; } 
            set 
            {
                if (value > 0 ) { amount = value; } 
                else { Console.WriteLine("Кол-во не может быть 0 или отрицательным"); }

            } 
        }

        public string ToString() 
        {
            return $"{num} {name} {type} {amount}{unit}";
        }

    }
}
