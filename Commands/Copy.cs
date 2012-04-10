using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class Copy : TransformColumnCommand, ITransformCommand
    {
        public string DestinationColumn;

        private object destinationColumnIdentifier;

        /// <summary>
        /// Если назначение не указано, то используем первый свободный столбец.
        /// </summary>
        [XmlIgnore]
        public bool DestinationColumnSpecified;

        public string ColumnHeader;

        /// <summary>
        /// Если не указан новый заголовок, то он остаётся пустым. Дублировать названия нельзя.
        /// </summary>
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

        protected override string ProcessCell(uint row, string cellValue)
        {
            if (!DestinationColumnSpecified)
            {
                uint headerRowIndex = GetStartRowBound();
                uint lastRowIndex = GetEndRowBound(headerRowIndex);
                destinationColumnIdentifier = GetLastColumnIndex(headerRowIndex, lastRowIndex) + 1;
            }

            Range newCell = worksheet.Cells[row, destinationColumnIdentifier];
            newCell.Value = cellValue;
            return cellValue;
        }

        /// <summary>
        /// Вычисляет новые границы данных.
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
                this.endColumnBound = destinationColumnNumber;
            }
        }
    }
}