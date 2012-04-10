using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class SplitByLength : TransformColumnCommand, ITransformCommand
    {
        /// <summary>
        /// Если длина отрицательна, то обработка идёт справа налево.
        /// </summary>
        public int Length;

        public string DestinationFirstColumn;

        private object destinationFirstColumnIdentifier;

        /// <summary>
        /// Если назначение не указано, то используем первые свободные столбцы
        /// </summary>
        [XmlIgnore]
        public bool DestinationFirstColumnSpecified;

        public string Column1Header;

        [XmlIgnore]
        public bool Column1HeaderSpecified;

        public string Column2Header;

        [XmlIgnore]
        public bool Column2HeaderSpecified;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);
            if (CheckDeserialization())
            {
                FindBoundsAndProcess();
            }

            return this;
        }

        #endregion


        /// <summary>
        /// Метод базового класса проверяет исходную колонку.
        /// Перегружнный метод проверяет новую колонку.
        /// </summary>
        /// <returns></returns>
        protected override bool CheckDeserialization()
        {
            if (base.CheckDeserialization())
            {
                if (!DestinationFirstColumnSpecified)
                    return true;

                uint number = 0;

                if (uint.TryParse(DestinationFirstColumn, out number))
                {
                    destinationFirstColumnIdentifier = number;
                    return true;
                }
                else
                {
                    Regex regex = new Regex("[a-zA-Z]+");
                    Match match = regex.Match(DestinationFirstColumn);

                    if (match.Success)
                    {
                        destinationFirstColumnIdentifier = DestinationFirstColumn;
                        return true;
                    }

                    return false;
                }
            }

            return false;
        }

        protected override void ProcessRows()
        {
            string cellValue = null;
            string firstPart = null;
            string secondPart = null;

            if (DestinationFirstColumnSpecified)
            {
                ComputeNewBounds();

                for (uint i = this.startRowBound.Value + 1; i <= endRowBound; i++)
                {
                    Range cell = worksheet.Cells[i, columnIdentifier];
                    cellValue = cell.Value == null ? "" : cell.Text.ToString();

                    GetParts(cellValue, ref firstPart, ref secondPart);

                    Range newCell1 = worksheet.Cells[i, destinationFirstColumnIdentifier];
                    Range newCell2 = worksheet.Cells[i, GetNextColumn(destinationFirstColumnIdentifier)];

                    newCell1.Value = firstPart;
                    newCell2.Value = secondPart;
                }

                if (Column1HeaderSpecified)
                    worksheet.Cells[startRowBound, destinationFirstColumnIdentifier].Value = Column1Header;

                if (Column2HeaderSpecified)
                    worksheet.Cells[startRowBound, GetNextColumn(destinationFirstColumnIdentifier)].Value = Column2Header;
            }
            else // если используются первые свободные столбцы
            {
                uint headerRowIndex = GetStartRowBound();
                uint lastRowIndex = GetEndRowBound(headerRowIndex);
                uint column1Index = GetIndexForNewColumn(headerRowIndex, lastRowIndex);
                uint column2Index = GetIndexForNewColumn(headerRowIndex, lastRowIndex);

                for (uint i = this.startRowBound.Value + 1; i <= endRowBound; i++)
                {
                    Range cell = worksheet.Cells[i, columnIdentifier];
                    cellValue = cell.Value == null ? "" : cell.Text.ToString();

                    GetParts(cellValue, ref firstPart, ref secondPart);

                    Range newCell1 = worksheet.Cells[i, column1Index];
                    Range newCell2 = worksheet.Cells[i, column2Index];

                    newCell1.Value = firstPart;
                    newCell2.Value = secondPart;
                }

                if (Column1HeaderSpecified)
                    worksheet.Cells[startRowBound, column1Index].Value = Column1Header;

                if (Column2HeaderSpecified)
                    worksheet.Cells[startRowBound, column2Index].Value = Column2Header;
            }
        }

        /// <summary>
        /// Вычисляет новые границы данных в случае, если столбцы назначения указаны явно.
        /// </summary>
        private void ComputeNewBounds()
        {
            uint destinationFirstColumnNumber = 0;

            if (destinationFirstColumnIdentifier.GetType() == typeof(uint))
            {
                destinationFirstColumnNumber = (uint)destinationFirstColumnIdentifier;
            }
            else
            {
                string destinationFirstColumnString = (string)destinationFirstColumnIdentifier;
                destinationFirstColumnNumber = Base26Decode(destinationFirstColumnString);
            }

            if (destinationFirstColumnNumber + 1 > this.endColumnBound)
            {
                this.endColumnBound = destinationFirstColumnNumber + 1;
            }
        }

        private void GetParts(string cellValue, ref string firstPart, ref string secondPart)
        {
            if (Length >= 0)
            {
                firstPart = cellValue.Substring(0, Math.Min(Length, cellValue.Length));
                secondPart = cellValue.Substring(Math.Min(Length, cellValue.Length));
            }
            else
            {
                // по сути вызов String.Length - ABS(Length)
                firstPart = cellValue.Substring(Math.Max(0, cellValue.Length + Length));
                secondPart = cellValue.Substring(0, Math.Max(0, cellValue.Length + Length));
            }
        }
    }
}
