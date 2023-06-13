using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon.Models
{
    public class Client
    {
        public int Code { get; set; }
        public string Name { get; set; } = "";
        public string Adress { get; set; } = "";
        public string ContactPerson { get; set; } = "";
        public Client() { }

        public Client(IXLRow row)
        {
            var cells = row.Cells().ToList();
            Code = (int)cells[0].Value.GetNumber();
            Name = cells[1].GetText();
            Adress = cells[2].GetText();
            ContactPerson = cells[3].GetText();
        }

        public override string ToString()
        {
            return $"код клиента: {Code},\n\tнаименование организации: {Name},\n\tадрес: {Adress},\n\tконтактное лицо(ФИО): {ContactPerson}";
        }
    }
}
