using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;

namespace XmlSerializationResearch
{
    public class A : II
    {
        public int i = 1;

        #region II Members

        public int getI()
        {
            return i;
        }

        #endregion
    }

    class B
    {
        public int i = 2;
    }

    public interface II
    {
        int getI();
    }

    class Program
    {
        static void Main(string[] args)
        {
            //AddColumn add = new XmlSerializationResearch.AddColumn();

            //XmlElementAttribute myElementAttribute = new XmlElementAttribute();
            //myElementAttribute.ElementName = "Source";
            //XmlAttributes myAttributes = new XmlAttributes();
            //myAttributes.XmlElements.Add(myElementAttribute);
            //XmlAttributeOverrides myOverrides = new XmlAttributeOverrides();
            //myOverrides.Add(typeof(TestCommand), "Column", myAttributes);
            //XmlSerializer serializer = new XmlSerializer(typeof(TestCommand), myOverrides);

            //MemoryStream stream = new MemoryStream();
            //serializer.Serialize(stream, add);
            //stream.Position = 0;

            //Console.WriteLine(new StreamReader(stream).ReadToEnd());

            //string a = "<TestCommand><Source>D</Source></TestCommand>";
            //TestCommand b = (TestCommand)serializer.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(a)));
            //b.SetExcelApp(new Microsoft.Office.Interop.Excel.Application());
            //b.SetFilePath(Path.Combine(System.Windows.Forms.Application.StartupPath, "Sheet.xls"));
            //b.CastColumnIdentifier();
            //b.Execute(null);
            //string r = "<RemoveColumn><ColumnIndex>1</ColumnIndex><ColumnName>1</ColumnName></RemoveColumn>";
            //RemoveColumn re = (RemoveColumn)serializer.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(r)));

            //Console.WriteLine("Done");
            //Console.ReadKey();
            //return;

            //string astr = "<A><i>3</i></A>";
            //II a = (II)ser1.Deserialize(new MemoryStream(Encoding.UTF8.GetBytes(astr)));

            //Console.WriteLine(a.getI());

            #region Main
            ScenarioReader reader = new ScenarioReader();
            IList<ITransformCommand> commands = reader.LoadCommands(Path.Combine(System.Windows.Forms.Application.StartupPath, @"TestCommands\Commands.xml"));
            if (commands.Count != 0)
            {
                commands[0].SetExcelApp(new Microsoft.Office.Interop.Excel.Application());
                commands[0].SetFilePath(Path.Combine(System.Windows.Forms.Application.StartupPath, "Sheet.xls"));
            }

            TransformCommand prevCommand = null;
            for (int i = 0; i < commands.Count; i++)
            {
                Console.WriteLine("{0}: Executing {1}", DateTime.Now, commands[i].GetType());
                prevCommand = commands[i].Execute(prevCommand);
            }
            #endregion

            //RenameHeader header = new RenameHeader();
            //header.OldName = new List<string>();
            //header.OldName.Add("FIO");
            //header.OldName.Add("ФИО");
            //header.NewName = "Column1";

            //XmlSerializer serializer = new XmlSerializer(typeof(RenameHeader));
            //MemoryStream stream = new MemoryStream();
            //serializer.Serialize(stream, header);
            //stream.Position = 0;

            //Console.WriteLine(new StreamReader(stream).ReadToEnd());

            ////string serializedHeader = "<RenameHeader><OldName><Item>FIO</Item><Item>ФИО</Item></OldName><NewName>Column2</NewName></RenameHeader>";
            //stream.Position = 0;
            //string serializedHeader = new StreamReader(stream).ReadToEnd();

            //MemoryStream streamToDeserialize = new MemoryStream(Encoding.UTF8.GetBytes(serializedHeader));
            //ITransformCommand deserialized = (ITransformCommand)serializer.Deserialize(streamToDeserialize);
            //deserialized.Excel = new Microsoft.Office.Interop.Excel.Application();
            //deserialized.FilePath = Path.Combine(System.Windows.Forms.Application.StartupPath, "Sheet.xls");


            //deserialized.SetExcelApp(new Microsoft.Office.Interop.Excel.Application());
            //deserialized.SetFilePath(Path.Combine(System.Windows.Forms.Application.StartupPath, "Sheet.xls"));
            //deserialized.Execute(null);

            //Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Visible = true;
            //// нужен полный путь к файлу
            //Workbook workbook = excel.Workbooks.Open(Path.Combine(System.Windows.Forms.Application.StartupPath, "Sheet.xls"));
            //Sheets sheets = workbook.Worksheets;
            //Worksheet worksheet = (Worksheet)sheets.get_Item(1);

            //Stopwatch stopwatch = new Stopwatch();
            //stopwatch.Start();

            //RemoveTopEmptyRows(excel, worksheet);
            //RemoveLeftEmptyColumns(excel, worksheet);
            //RenameHeader(excel, worksheet, "column5", "abababa");
            //RemoveRow(worksheet, 5);
            //RemoveRow(worksheet, 3);
            ////RemoveColumn(worksheet, 2);
            ////RemoveColumn(worksheet, "B");

            //ushort firstRow = GetFirstNonEmptyRowHeader(excel, worksheet);
            //ushort lastRow = GetLastNonEmptyRowForward(excel, worksheet, firstRow);
            //Console.WriteLine(firstRow);
            //Console.WriteLine(lastRow);
            //Console.WriteLine(GetIndexForNewColumn(excel, worksheet, firstRow, lastRow));

            //AddColumn(excel, worksheet, "newColumn1", "newValue");
            //AddColumn(excel, worksheet, "newColumn2", DateTime.Now);
            ////AddColumn(excel, worksheet, "newColumn3", DateTime.Now, "-1d2h3m8s");
            ////AddColumn(excel, worksheet, "newColumn5", 1, 3);
            ////AddColumn(excel, worksheet, "newColumn6", 1.5, 3.5);
            ////AddColumn(excel, worksheet, "newColumn7", 1.5);

            //string[] values = null;


            //stopwatch.Stop();

            //Console.WriteLine("Excel process time: " + stopwatch.Elapsed);
            //Console.WriteLine("");
            Console.WriteLine("Done.");
            Console.ReadKey();
        }

        //private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet, string columnHeaderName, double value)
        //{
        //    throw new NotImplementedException();
        //}

        private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel,
                                    Worksheet worksheet,
                                    string columnHeaderName,
                                    double startValue,
                                    double step)
        {
            throw new NotImplementedException();
        }

        private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel,
                                    Worksheet worksheet,
                                    string columnHeaderName,
                                    int startvalue,
                                    int step)
        {
            throw new NotImplementedException();
        }

        private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel,
                                    Worksheet worksheet,
                                    string columnHeaderName,
                                    DateTime startValue,
                                    string stepPattern)
        {
            throw new NotImplementedException();
        }

        //private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel,
        //                                Worksheet worksheet,
        //                                string columnHeaderName,
        //                                DateTime value)
        //{
        //    uint headerRowIndex = GetFirstNonEmptyRowHeader(excel, worksheet);
        //    uint lastRowIndex = GetLastNonEmptyRowForward(excel, worksheet, 1);
        //    int columnIndex = GetIndexForNewColumn(excel, worksheet, headerRowIndex, lastRowIndex);
        //    Range headerCell = worksheet.Cells[headerRowIndex, columnIndex];
        //    headerCell.Value = columnHeaderName;

        //    Range cellStart = worksheet.Cells[headerRowIndex + 1, columnIndex];
        //    Range cellEnd = worksheet.Cells[lastRowIndex, columnIndex];
        //    Range cells = worksheet.Range[cellStart, cellEnd];
        //    cells.Value = value.ToShortDateString();
        //}

        ///// <summary>
        ///// Добавление одного текстового значения во все ячейки.
        ///// Столбец добавляется после всех столбцов с данными.
        ///// </summary>
        ///// <param name="excel"></param>
        ///// <param name="worksheet"></param>
        ///// <param name="columnHeaderName">Заголовок столбца.</param>
        ///// <param name="value">Значение, помещаемое в каждую ячейку столбца.</param>
        //private static void AddColumn(Microsoft.Office.Interop.Excel.Application excel,
        //                                Worksheet worksheet,
        //                                string columnHeaderName,
        //                                string value)
        //{
        //    uint headerRowIndex = GetFirstNonEmptyRowHeader(excel, worksheet);
        //    uint lastRowIndex = GetLastNonEmptyRowForward(excel, worksheet, 1);
        //    int columnIndex = GetIndexForNewColumn(excel, worksheet, headerRowIndex, lastRowIndex);
        //    Range headerCell = worksheet.Cells[headerRowIndex, columnIndex];
        //    headerCell.Value = columnHeaderName;

        //    Range cellStart = worksheet.Cells[headerRowIndex + 1, columnIndex];
        //    Range cellEnd = worksheet.Cells[lastRowIndex, columnIndex];
        //    Range cells = worksheet.Range[cellStart, cellEnd];
        //    cells.Value = value;
        //}

        private static int GetIndexForNewColumn(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet, uint headerRowIndex, uint lastRowIndex)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            int columnIndex = 256;

            while (excel.WorksheetFunction.CountA(GetColumnRange(worksheet, headerRowIndex, lastRowIndex, columnIndex)) == 0)
            {
                columnIndex--;
            }

            stopwatch.Stop();
            Console.WriteLine(stopwatch.Elapsed);
            return columnIndex + 1;
        }

        private static Range GetColumnRange(Worksheet worksheet, uint headerRowIndex, uint lastRowIndex, int columnIndex)
        {
            Range range1 = worksheet.Cells[headerRowIndex, columnIndex];
            Range range2 = worksheet.Cells[lastRowIndex, columnIndex];
            Range cells = worksheet.Range[range1, range2];

            return cells;
        }

        //private static void RemoveRow(Worksheet worksheet, uint rowIndex)
        //{
        //    worksheet.get_Range("A" + rowIndex.ToString()).EntireRow.Delete();
        //}

        //private static void RemoveColumn(Worksheet worksheet, uint columnIndex)
        //{
        //    Range column = worksheet.Columns[columnIndex];
        //    column.EntireColumn.Delete();
        //}

        //private static void RemoveColumn(Worksheet worksheet, string columnName)
        //{
        //    Range column = worksheet.Columns[columnName];
        //    column.EntireColumn.Delete();
        //}


        //private static void RenameHeader(Microsoft.Office.Interop.Excel.Application excel,
        //                                Worksheet worksheet,
        //                                string oldName,
        //                                string newName)
        //{
        //    uint row = GetFirstNonEmptyRowHeader(excel, worksheet);

        //    Array values = (Array)worksheet.get_Range("A" + row.ToString()).EntireRow.Cells.Value;
        //    int column = GetColumnIndex(values, oldName);

        //    if (column != 0)
        //    {
        //        Range cell = worksheet.Cells[row, column];
        //        cell.Value = newName;
        //    }
        //}

        private static ushort GetFirstNonEmptyRowHeader(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet)
        {
            // определяем первую непустую строку
            ushort row = 1;

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A" + row.ToString(), "A" + row.ToString()).EntireRow) == 0)
            {
                row++;
            }
            return row;
        }

        private static ushort GetLastNonEmptyRowBackward(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet)
        {
            // определяем последнюю непустую строку
            ushort row = ushort.MaxValue;

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A" + row.ToString(), "A" + row.ToString()).EntireRow) == 0)
            {
                row--;
            }

            return row;
        }

        private static ushort GetLastNonEmptyRowForward(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet, ushort startIndex)
        {
            ushort row = startIndex;

            while (excel.WorksheetFunction.CountA(worksheet.get_Range("A" + row.ToString(), "A" + row.ToString()).EntireRow) != 0)
            {
                row++;
            }

            return (ushort)(row - 1);
        }

        //private static int GetColumnIndex(Array values, string oldName)
        //{
        //    for (int i = 1; i < values.Length; i++)
        //    {
        //        if (values.GetValue(1, i) != null)
        //            if (values.GetValue(1, i).ToString() == oldName)
        //                return i;
        //    }

        //    // не нашли
        //    return 0;
        //}

        //private static void RemoveTopEmptyRows(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet)
        //{
        //    while (excel.WorksheetFunction.CountA(worksheet.get_Range("A1", "A1").EntireRow) == 0)
        //    {
        //        worksheet.get_Range("A1", "A1").EntireRow.Delete();
        //    }
        //}

        private static void RemoveEmptyRows(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet, ushort maxEmptyRowsArea)
        {
            throw new NotImplementedException("Лениво писать её :(");
            //ushort continuousDeletedRows = 0;
            //ushort currentRow = 1;

            //while (excel.WorksheetFunction.CountA(worksheet.get_Range("A1", "A1").EntireRow) == 0)
            //{
            //    worksheet.get_Range("A1", "A1").EntireRow.Delete();
            //    continuousDeletedRows++;

            //    if (continuousDeletedRows > maxEmptyRowsArea)
            //        return;
            //}
        }

        //private static void RemoveLeftEmptyColumns(Microsoft.Office.Interop.Excel.Application excel, Worksheet worksheet)
        //{
        //    while (excel.WorksheetFunction.CountA(worksheet.get_Range("A1", "A1").EntireColumn) == 0)
        //    {
        //        worksheet.get_Range("A1", "A1").EntireColumn.Delete();
        //    }
        //}

        private static string[] RowToArray(string[] values, Array myvalues)
        {
            values = new string[myvalues.Length];
            for (int j = 1; j <= values.Length; j++)
            {
                object value = myvalues.GetValue(1, j);
                if (value == null)
                    values[j - 1] = "";
                else
                    values[j - 1] = value.ToString();
            }
            return values;
        }
    }
}
