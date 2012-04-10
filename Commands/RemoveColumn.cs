using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Xml.Serialization;

namespace XmlSerializationResearch
{
    public class RemoveColumn : TransformColumnCommand, ITransformCommand
    {
        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);
            CheckDeserialization();

            Range column = worksheet.Columns[columnIdentifier];
            column.EntireColumn.Delete();

            this.endColumnBound = null;

            return this;
        }

        #endregion
    }
}
