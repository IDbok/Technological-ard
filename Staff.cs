﻿
namespace Technological_card
{
    internal class Staff //1. Требования к составу бригады и квалификации
    {
        int num;
        string name;
        string type;
        string combineResponsibility;
        byte elSaftyGroup;
        byte grade;
        string competence;
        string comment;

        public Staff(int num, string name, string type, string combineResponsibility, 
            byte elSaftyGroup, byte grade, string competence, string comment) 
        {
            this.num = num;
            this.name = name;
            this.type = type;
            this.combineResponsibility = combineResponsibility;
            this.elSaftyGroup = elSaftyGroup;
            this.grade = grade;
            this.competence = competence;
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
        public byte ElSaftyGroup
        { get { return elSaftyGroup; } set { elSaftyGroup = value; } }
        public byte Grade
        { get { return grade; } set {  grade = value; } }
        public string Competence
        { get { return competence; } set { competence = value; } }
        public string Comment
        { get { return comment; } set {  comment = value; } }
    }
}
