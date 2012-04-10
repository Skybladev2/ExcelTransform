using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlSerializationResearch
{
    public class RemoveLeftEmptyColumns : TransformCommand, ITransformCommand
    {
        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A1", "A1").EntireColumn) == 0)
            {
                worksheet.get_Range("A1", "A1").EntireColumn.Delete();

                this.startColumnBound = 1;
                this.endColumnBound = null;
            }

            return this;
        }

        #endregion
    }
}
