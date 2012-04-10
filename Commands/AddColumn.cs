using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace XmlSerializationResearch
{
    public class AddColumn : TransformCommand, ITransformCommand
    {
        public string ColumnHeaderName;

        /// <summary>
        /// Если не null - заполняем ячейки данной строкой.
        /// </summary>
        public string Value;

        /// <summary>
        /// Шаблон для инкремента/декремента даты-времени
        /// </summary>
        [XmlAttribute()]
        public string stepPattern;

        [XmlAttribute()]
        public double step;

        [XmlIgnoreAttribute]
        public bool stepSpecified;

        [XmlAttribute()]
        public string dataType;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            uint headerRowIndex = GetStartRowBound();
            uint lastRowIndex = GetEndRowBound(headerRowIndex);
            uint columnIndex = GetIndexForNewColumn(headerRowIndex, lastRowIndex);

            Range headerCell = worksheet.Cells[headerRowIndex, columnIndex];
            headerCell.Value = this.ColumnHeaderName;

            Range cellStart = worksheet.Cells[headerRowIndex + 1, columnIndex];
            Range cellEnd = worksheet.Cells[lastRowIndex, columnIndex];
            Range cells = worksheet.Range[cellStart, cellEnd];

            double doubleStartValue = 0;
            int intStartValue = 0;
            int intStep = 0;
            DateTime dateStartValue = DateTime.MinValue;
            string localDataType = null;

            if (!string.IsNullOrEmpty(dataType))
            {
                localDataType = dataType.ToLowerInvariant();

                switch (localDataType)
                {
                    case "double":
                        NumberFormatInfo nfi = new NumberFormatInfo();
                        nfi.NumberDecimalSeparator = ".";
                        doubleStartValue = Double.Parse(Value, nfi);
                        break;
                    case "int":
                        intStartValue = Int32.Parse(Value, NumberStyles.Integer);
                        intStep = (int)step;
                        break;
                    case "datetime":
                        dateStartValue = DateTime.Parse(Value);
                        break;
                    default:
                        throw new ArgumentException("Неизвестный тип данных", Value);
                        break;
                }
            }

            // если тип данных указан - значит будет изменение значения
            if (!string.IsNullOrEmpty(this.dataType))
            {
                int counter = 0;
                for (uint i = headerRowIndex + 1; i <= lastRowIndex; i++)
                {
                    Range cell = worksheet.Cells[i, columnIndex];

                    switch (localDataType)
                    {
                        case "double":
                            cell.Value = doubleStartValue + step * counter++;
                            break;
                        case "int":
                            cell.Value = intStartValue + intStep * counter++;
                            break;
                        case "datetime":
                            cell.Value = IncrementDate(stepPattern, dateStartValue, counter++);
                            break;
                        default:
                            break;
                    }
                }
            }
            else
                cells.Value = this.Value;

            return this;
        }

        private DateTime IncrementDate(string stepPattern, DateTime dateStartValue, int multiplier)
        {
            DateTime newDateTime = dateStartValue;

            // 1y2M3d4h5m6s
            Regex regex = new Regex("(\\+|-)*(((?<year>\\d+)y)?((?<month>\\d+)M)?((?<day>\\d+)d)?((?<hour>\\d+)h)?((?<minute>\\d+)m)?((?<second>\\d+)s)?)+");
            Match match = regex.Match(stepPattern);
            int sign = stepPattern[0] == '-' ? -1 : 1;

            if (match.Success)
            {
                GroupCollection groups = match.Groups;
                if (!string.IsNullOrEmpty(groups["year"].Value))
                    newDateTime = newDateTime.AddYears(int.Parse(groups["year"].Value) * sign * multiplier);
                if (!string.IsNullOrEmpty(groups["month"].Value))
                    newDateTime = newDateTime.AddMonths(int.Parse(groups["month"].Value) * sign * multiplier);
                if (!string.IsNullOrEmpty(groups["day"].Value))
                    newDateTime = newDateTime.AddDays(int.Parse(groups["day"].Value) * sign * multiplier);
                if (!string.IsNullOrEmpty(groups["hour"].Value))
                    newDateTime = newDateTime.AddDays(int.Parse(groups["hour"].Value) * sign * multiplier);
                if (!string.IsNullOrEmpty(groups["minute"].Value))
                    newDateTime = newDateTime.AddDays(int.Parse(groups["minute"].Value) * sign * multiplier);
                if (!string.IsNullOrEmpty(groups["second"].Value))
                    newDateTime = newDateTime.AddDays(int.Parse(groups["second"].Value) * sign * multiplier);
            }

            return newDateTime;
        }

        #endregion
    }
}