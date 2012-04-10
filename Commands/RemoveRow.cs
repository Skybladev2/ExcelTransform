using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlSerializationResearch
{
    public class RemoveRow : TransformCommand, ITransformCommand
    {
        public uint RowIndex;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            worksheet.get_Range("A" + RowIndex.ToString()).EntireRow.Delete();

            // так как мы не знаем, что идёт после удаляемой строки,
            // то не можем определить, какой строкой кончаются данные
            this.endRowBound = null;

            return this;
        }

        #endregion
    }
}
