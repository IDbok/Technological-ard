
namespace Technological_card
{
    internal class Staff //1. Требования к составу бригады и квалификации
    {
        int num;
        string name;
        string type;
        string combineResponsibility;
        string elSaftyGroup;
        byte grade;
        string competence;
        string symbol;
        string comment;

        public Staff()
        {

        }
        public Staff(int num, string name, string type, string combineResponsibility,
            string elSaftyGroup, byte grade, string competence, string symbol, string comment) 
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.combineResponsibility = combineResponsibility;
            this.elSaftyGroup = elSaftyGroup;
            this.grade = grade;
            this.competence = competence;
            this.symbol = symbol;
            this.comment = comment;
        }


        public int Num 
        { 
            get { return num; }
            set { num = value; }
        }

        public string Name 
        { 
            get { return name; }
            set { name = value; }
        }
        public string Type 
        {
            get { return type; }
            set { type = value; }
        }
        public string CombineResponsibility
        {
            get { return combineResponsibility; }
            set { combineResponsibility = value; }
        }
        public string ElSaftyGroup
        { get { return elSaftyGroup; } set { elSaftyGroup = value; } }
        public byte Grade
        { get { return grade; } set {  grade = value; } }
        public string Competence
        { get { return competence; } set { competence = value; } }
        public string Symbol
        { get { return symbol; } set { symbol = value; } }
        public string Comment
        { get { return comment; } set {  comment = value; } }

        public string ToSring() 
        {
            return $"{num} {name} {type} {combineResponsibility} {elSaftyGroup} {grade} {competence} {symbol}" +
                $"\n{comment}";
        }
    }
}
