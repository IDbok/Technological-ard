namespace Technological_card
{
    internal class Tool : Struct //5. Требования к инструментам и приспособлениям
    {
        int num;
        string name;
        string type;
        string unit;
        int amount;

        public Tool (int num, string name, string type, string unit, int amount)
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.unit = unit;
            this.amount = amount;
        }

        public override int Num
        {
            get { return num; }
            set { num = value; }
        }

        public override string Name
        {
            get { return name; }
            set { name = value; }
        }
        public override string Type
        {
            get { return type; }
            set { type = value; }
        }
        
        public override string Unit
        { get { return unit; } set { unit = value; } }
        public override int Amount
        { get { return amount; } set { amount = value; } }
    }
}
