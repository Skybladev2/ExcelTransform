using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class MergeColumns : TransformRangeCommand, ITransformCommand
    {
        public string Delimiter;

        public string DestinationColumn;

        private object destinationColumnIdentifier;

        /// <summary>
        /// Если назначение не указано, то записываем результат в первый свободный столбец.
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

        protected override void FindBoundsAndProcess()
        {
            uint headerRowIndex = GetStartRowBound();
            uint lastRowIndex = GetEndRowBound(headerRowIndex);

            if (!DestinationColumnSpecified)
                destinationColumnIdentifier = GetLastColumnIndex(headerRowIndex, lastRowIndex) + 1;

            // если указаны колонки - обрабатываем их, иначе все ячейки с данными
            if (SourceSpecified)
                for (int i = 0; i < sourceColumnIdentifiers.Count; i++)
                {
                    object columnIdentifier = sourceColumnIdentifiers[i];

                    for (uint j = this.startRowBound.Value + 1; j <= endRowBound; j++)
                    {
                        Range cell = worksheet.Cells[j, columnIdentifier];
                        string cellValue = cell.Value == null ? "" : cell.Text.ToString();

                        Range newCell = worksheet.Cells[j, destinationColumnIdentifier];

                        if (i == 0)
                            newCell.Value = cellValue;
                        else
                            newCell.Value += Delimiter + cellValue;
                    }
                }
            else
            {
                // надеемся, что какие-нибудь данные всё же есть :)
                GetLastColumnIndex(headerRowIndex, lastRowIndex);
                GetFirstNonEmptyColumn();

                for (uint i = this.startColumnBound.Value; i < this.endColumnBound; i++)
                {
                    object columnIdentifier = i;

                    for (uint j = this.startRowBound.Value + 1; j <= endRowBound; j++)
                    {
                        Range cell = worksheet.Cells[j, columnIdentifier];
                        string cellValue = cell.Value == null ? "" : cell.Text.ToString();

                        Range newCell = worksheet.Cells[j, destinationColumnIdentifier];

                        if (i == this.startColumnBound)
                            newCell.Value = cellValue;
                        else
                            newCell.Value += Delimiter + cellValue;
                    }
                }
            }

            if (ColumnHeaderSpecified)
            {
                Range cell = worksheet.Cells[headerRowIndex, destinationColumnIdentifier];
                cell.Value = ColumnHeader;
            }

            ComputeNewBounds();
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
