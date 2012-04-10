using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    /// <summary>
    /// Получает подстроку.
    /// Может записывать результат в тот же столбец либо в новый.
    /// Если указывается отрицательный начальный индекс, то подстрока выбирается с правого края.
    /// Если не указывается столбец назначения, то результат записывается в тот же столбец.
    /// </summary>
    public class Substring : TransformColumnCommand, ITransformCommand
    {
        public int StartIndex;

        public int Length;

        public string DestinationColumn;

        private object destinationColumnIdentifier;

        /// <summary>
        /// Если назначение не указано, то записываем результат в исходный столбец
        /// </summary>
        [XmlIgnore]
        public bool DestinationColumnSpecified;

        [XmlIgnore]
        public bool StartIndexSpecified;

        [XmlIgnore]
        public bool LengthSpecified;

        public string ColumnHeader;

        [XmlIgnore]
        public bool ColumnHeaderSpecified;

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

        protected override void FindBoundsAndProcess()
        {
            base.FindBoundsAndProcess();

            if (ColumnHeaderSpecified)
            {
                uint headerRowIndex = GetStartRowBound();
                Range cell = worksheet.Cells[headerRowIndex, destinationColumnIdentifier];
                cell.Value = ColumnHeader;
            }

            ComputeNewBounds();
        }

        /// <summary>
        /// Вычисляет новые границы данных в случае, если столбцы назначения указаны явно.
        /// </summary>
        private void ComputeNewBounds()
        {
            uint destinationColumnNumber = 0;

            if (destinationColumnIdentifier.GetType() == typeof(uint))
            {
                destinationColumnNumber = (uint)destinationColumnIdentifier;
            }
            else
            {
                string destinationFirstColumnString = (string)destinationColumnIdentifier;
                destinationColumnNumber = Base26Decode(destinationFirstColumnString);
            }

            if (destinationColumnNumber > this.endColumnBound)
            {
                this.endColumnBound = destinationColumnNumber + 1;
            }
        }

        /// <summary>
        /// Метод базового класса проверяет исходную колонку.
        /// Перегруженный метод проверяет новую колонку.
        /// </summary>
        /// <returns></returns>
        protected override bool CheckDeserialization()
        {
            if (base.CheckDeserialization())
            {
                if (!DestinationColumnSpecified)
                    return true;

                uint number = 0;

                if (uint.TryParse(DestinationColumn, out number))
                {
                    destinationColumnIdentifier = number;
                    return true;
                }
                else
                {
                    Regex regex = new Regex("[a-zA-Z]+");
                    Match match = regex.Match(DestinationColumn);

                    if (match.Success)
                    {
                        destinationColumnIdentifier = DestinationColumn;
                        return true;
                    }

                    return false;
                }
            }

            return false;
        }

        protected override void ProcessRows()
        {
            if (!DestinationColumnSpecified)
                GetBounds(); // используется первый свободный столбец

            ProcessRowsCycle();

            if (ColumnHeaderSpecified)
                worksheet.Cells[startRowBound, destinationColumnIdentifier].Value = ColumnHeader;
        }

        private void ProcessRowsCycle()
        {
            string cellValue = null;
            string part = null;

            for (uint i = this.startRowBound.Value + 1; i <= endRowBound; i++)
            {
                Range cell = worksheet.Cells[i, columnIdentifier];
                cellValue = cell.Value == null ? "" : cell.Text.ToString();

                part = GetPart(cellValue);

                Range newCell = worksheet.Cells[i, destinationColumnIdentifier];

                newCell.Value = part;
            }
        }

        private void GetBounds()
        {
            uint headerRowIndex = GetStartRowBound();
            uint lastRowIndex = GetEndRowBound(headerRowIndex);
            destinationColumnIdentifier = GetIndexForNewColumn(headerRowIndex, lastRowIndex);
        }

        private string GetPart(string cellValue)
        {
            if (StartIndexSpecified)
                return GetPartWithStartIndex(cellValue);
            else
                return GetPartWithoutStartIndex(cellValue);
        }

        private string GetPartWithoutStartIndex(string cellValue)
        {
            if (LengthSpecified)
            {
                return GetSubstringWithLengthOnly(cellValue);
            }
            else
            {
                return GetPartWithoutStartIndexAndLength(cellValue);
            }
        }

        /// <summary>
        /// Ничего не делаем.
        /// </summary>
        /// <param name="cellValue"></param>
        /// <returns>Исходная строка.</returns>
        private string GetPartWithoutStartIndexAndLength(string cellValue)
        {
            return cellValue;
        }

        /// <summary>
        /// Берём начало строки (или конец, если длина отрицательная) указанной длины.
        /// </summary>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private string GetSubstringWithLengthOnly(string cellValue)
        {
            if (Length >= 0)
                return cellValue.Substring(0, Math.Min(Length, cellValue.Length));
            else
                return cellValue.Substring(Math.Max(0, cellValue.Length + Length));
        }

        private string GetPartWithStartIndex(string cellValue)
        {
            if (LengthSpecified)
            {
                return GetSubstringWithStartIndexAndLength(cellValue);
            }
            else
            {
                return GetSubstringWithStartIndexOnly(cellValue);
            }
        }

        /// <summary>
        /// Берём конец строки (или начало, если индекс отрицательный), начиная с указанного символа.
        /// </summary>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private string GetSubstringWithStartIndexOnly(string cellValue)
        {
            if (StartIndex >= 0)
                return cellValue.Substring(Math.Min(StartIndex, cellValue.Length));
            else
                return cellValue.Substring(0, Math.Min(0, cellValue.Length + StartIndex));
        }

        /// <summary>
        /// Ультимативная функция. printf плачет и падает на колени перед ней.
        /// Начало извлечения и направление определяются знаками аргументов.
        /// Если StartIndex неотрицательный, то извлечение начинается с начала строки, если отрицательный - с конца.
        /// Если Length неотрицательная, то извлекаются символы справа от StartIndex, если отрицательная - слева.
        /// </summary>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        private string GetSubstringWithStartIndexAndLength(string cellValue)
        {
            int maxLength = 0;
            //if (Math.Abs(StartIndex) + Math.Abs(Length) > cellValue.Length)
            //    return cellValue;

            if (StartIndex >= 0)
            {
                maxLength = cellValue.Length - Math.Abs(StartIndex);

                if (Length >= 0)
                    return cellValue.Substring(Math.Min(cellValue.Length, StartIndex), Math.Min(maxLength, Length));
                else
                    return cellValue.Substring(Math.Min(cellValue.Length, StartIndex + Length), Math.Min(maxLength, -Length));
            }
            else
            {
                maxLength = Math.Abs(StartIndex);

                if (Length >= 0)
                    return cellValue.Substring(Math.Max(0, cellValue.Length + StartIndex), Math.Min(maxLength, Length));
                else
                    return cellValue.Substring(Math.Max(0, cellValue.Length + StartIndex + Length), Math.Min(maxLength, -Length));
            }
        }
    }
}
