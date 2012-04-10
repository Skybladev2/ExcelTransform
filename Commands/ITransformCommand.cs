using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace XmlSerializationResearch
{
    public interface ITransformCommand
    {
        TransformCommand Execute(TransformCommand previousCommand);
        void SetExcelApp(Microsoft.Office.Interop.Excel.Application excel);
        void SetFilePath(string filePath);
    }
}
