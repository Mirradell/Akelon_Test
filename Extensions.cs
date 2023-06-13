using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Akelon
{
    public static class Extensions
    {
        public static Measure MeasureFromText(this string text)
        {
            switch (text)
            {
                case "Литр": return Measure.Liter;
                case "Килограмм": return Measure.Kilogram;
                case "Штука": return Measure.Thing;
                default: throw new ArgumentException($"Неизвестная единица измерения {text}");
            }
        }

        public static string MonthFromInt(this int month)
        {
            switch (month)
            {
                case 1:
                    return "январь";
                case 2:
                    return "февраль";
                case 3:
                    return "март";
                case 4:
                    return "апрель";
                case 5:
                    return "май";
                case 6:
                    return "июнь";
                case 7:
                    return "июль";
                case 8:
                    return "август";
                case 9:
                    return "сентябрь";
                case 10:
                    return "октябрь";
                case 11:
                    return "ноябрь";
                case 12:
                    return "декабрь";
                default:
                    return "нет такого месяца";
            }
        }
    }
}
