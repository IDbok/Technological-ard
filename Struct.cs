using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Technological_card
{
    public class Struct
    {
        virtual public int Num { get; set; }

        virtual public string Name { get; set; }

        virtual public string Type { get; set; }

        virtual public string Unit { get; set; }
        virtual public int Amount { get; set; }
        //virtual public float Price { get; set; }
    }
}
