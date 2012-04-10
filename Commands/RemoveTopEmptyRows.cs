using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlSerializationResearch
{
    public class RemoveTopEmptyRows : TransformCommand, ITransformCommand
    {
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
