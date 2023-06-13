using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon.Models
{
    public class Ware
    {
        public int Code { get; set; }
        public string Name { get; set; } = "";
        public Measure MeasureUnit { get; set; }
        public double PricePerOne { get; set; }
        
        public Ware() { }
        public Ware(IXLRow row)
        {
            var cells = row.Cells().ToList();
            Code = (int)cells[0].Value.GetNumber();
            Name = cells[1].GetText();
            MeasureUnit = cells[2].GetText().MeasureFromText();
            PricePerOne = cells[3].Value.GetNumber();
        }

        public override string ToString()
        {
            return $"код товара: {Code},\n\tнаименование: {Name},\n\tединицы измерения: {MeasureUnit},\n\tцена за единицу: {PricePerOne} руб";
        }
    }
}
