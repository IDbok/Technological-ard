using CsvHelper.Configuration.Attributes;


namespace Technological_card
{
    internal class ComponentModel //2. Требования к материалам и комплектующим
    {
        static int count { get; set; }
        [Name("Num")]
        public int num { get; set; }
        [Name("Name1")]
        public string name { get; set; }
        [Name("Type")]
        public string type { get; set; }
        [Name("Unit")]
        public string unit { get; set; }
        [Name("Amount")]
        public int amount { get; set; } 
    }

}
