using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    public class ReplaceRegex : TransformRangeCommand, ITransformCommand
    {
        public string Pattern;

        public string Replacement;

        private Regex regex = null;

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
            regex = new Regex(Pattern);

            base.FindBoundsAndProcess();
        }

        protected override string ProcessCell(string cellValue)
        {
            if (cellValue == null)
                cellValue = "";

            Match match = regex.Match(cellValue);

            if (match.Success)
            {
                GroupCollection gc = match.Groups;
                for (int j = 0; j < gc.Count; j++)
                {
                    CaptureCollection cc = gc[j].Captures;

                    int counter = cc.Count;

                    for (int k = 0; k < counter; k++)
                    {
                        if (!string.IsNullOrEmpty(cc[k].Value))
                            cellValue = cellValue.Replace(cc[k].Value, Replacement);
                    }
                }
            }
            return cellValue;
        }
    }
}
