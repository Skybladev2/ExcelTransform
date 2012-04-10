using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace XmlSerializationResearch
{
    /// <summary>
    /// Добавляет заданный текст слева или справа от значения
    /// </summary>
    public class Concatenate : TransformColumnCommand, ITransformCommand
    {
        public string What;

        [XmlAttribute]
        public Side side;

        public enum Side
        {
            [XmlEnum(Name = "left")]
            Left,
            [XmlEnum(Name = "right")]
            Right
        }

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
            if (side == Side.Left)
                return What + cellValue;
            else
                return cellValue + What;
        }
    }
}
