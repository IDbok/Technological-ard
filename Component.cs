using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Technological_card
{
    internal class Component //2. Требования к материалам и комплектующим
    {
        int num;
        string name;
        string type;
        string unit;
        uint amount; // using uint here

        public Component (int num, string name, string type, string unit, uint amount)
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.unit = unit;
            this.amount = amount;
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
    }
}
