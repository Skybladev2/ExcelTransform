using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace XmlSerializationResearch
{
    public class SplitColumn : TransformColumnCommand, ITransformCommand
    {
        public List<string> DestinationColumns;

        [XmlIgnore]
        private bool DestinationColumnsSpecified;

        public string Delimiter;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A1", "A1").EntireRow) == 0)
            {
                worksheet.get_Range("A1", "A1").EntireRow.Delete();

                // неохота выносить присвоение за цикл, чтобы опять проверять условие
                this.startRowBound = 1;
                this.endRowBound = null;
            }

            return this;
        }

        #endregion
    }
}
