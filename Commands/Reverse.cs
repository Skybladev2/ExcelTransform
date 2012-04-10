using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlSerializationResearch
{
    public class Reverse : TransformColumnCommand, ITransformCommand
    {
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

        protected override string ProcessCell(uint row, string cellValue)
        {
            if (cellValue == null)
                return null;

            char[] rev = cellValue.ToCharArray();
            Array.Reverse(rev);
            return (new string(rev));
        }
    }
}
