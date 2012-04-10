using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Excel;

namespace XmlSerializationResearch
{
    /// <summary>
    /// Добавляет новой строкой над блоком данных заголовки столбцов.
    /// Если эта строка пустая - остальные строки не сдвигаются.
    /// Если вверху нет места, то добавляется пустая строка и в неё записываются названия заголовков.
    /// </summary>
    public class AddHeaders : TransformCommand, ITransformCommand
    {
        [XmlArrayItem("Name")]
        public List<string> ColumnHeaders;

        #region ITransformCommand Members

        public TransformCommand Execute(TransformCommand previousCommand)
        {
            Init(previousCommand);

            uint headerRowIndex = GetStartRowBound();
            ushort firstColumnIndex = GetFirstNonEmptyColumn();

            // есть пустые строки до начала данных
            if (headerRowIndex != 1)
            {
                headerRowIndex--;
                if (this.startRowBound != null)
                    startRowBound--;
            }
            else
            {
                Range headerCell = worksheet.Cells[1, 1];
                headerCell.EntireRow.Insert(XlInsertShiftDirection.xlShiftDown);
                if (this.endRowBound != null)
                    this.endRowBound++;
            }

            int counter = 0;
            for (ushort i = firstColumnIndex; i < firstColumnIndex + ColumnHeaders.Count; i++)
            {
                Range cell = worksheet.Cells[headerRowIndex, i];
                cell.Value = ColumnHeaders[counter++];
            }

            return this;
        }

        #endregion
    }
}
