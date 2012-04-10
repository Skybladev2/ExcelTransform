using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class Replace : TransformRangeCommand, ITransformCommand
    {
        public string What;

        public string With;

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

        protected override string ProcessCell(string cellValue)
        {
            if (cellValue == null)
                cellValue = "";

            return cellValue.Replace(What, With);
        }
    }
}
