using ClosedXML.Excel;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon.Models
{
    public class Request
    {
        public int RequestCode { get; set; }
        public int WareCode { get; set; }
        public int ClientCode { get; set; }
        public int RequestNumber { get; set; }
        public int Count { get; set; }
        public DateTime Date { get; set; }

        public Request() { }
        public Request(IXLRow row)
        {
            var cells = row.Cells().ToList();
            RequestCode = (int)cells[0].Value.GetNumber();
            WareCode = (int)cells[1].Value.GetNumber();
            ClientCode = (int)cells[2].Value.GetNumber();
            RequestNumber = (int)cells[3].Value.GetNumber();
            Count = (int)cells[4].Value.GetNumber();
            Date = cells[5].Value.GetDateTime();
        }

        public override string ToString()
        {
            return $"код заявки: {RequestCode},\n\tкод товара: {WareCode},\n\tкод клиента: {ClientCode}," +
                   $"\n\tномер заявки: {RequestNumber},\n\tтребуемое количество: {Count},\n\tдата размещения: {Date}";
        }
    }
}
