using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Technological_card
{
    internal class Machine //3. Требования к механизмам
    {
        int num;
        string name;
        string type;
        string unit;
        uint amount; // using uint here
        decimal price;

        public Machine(int num, string name, string type, string unit, uint amount, decimal price)
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.unit = unit;
            this.amount = amount;
            this.price = price;
        }

        public int Num
        { get { return num; } set { num = value; } }

        public string Name
        { get { return name; } set { name = value; } }
        public string Type
        { get { return type; } set { type = value; } }

        public string Unit
        { get { return unit; } set { unit = value; } }
        public uint Amount
        { get { return amount; } set { amount = value; } }
        public decimal Price
        { get { return price; } set { price = value; } }
    }
}
