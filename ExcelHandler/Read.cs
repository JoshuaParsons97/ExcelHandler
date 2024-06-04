using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace ExcelHandler
{
    public class Read
    {
        public static DataSet ReadToDataSet(string FilePath, bool ColumnNamesSpecified = true)
        {
            if (File.Exists(FilePath))
            {
                try
                {
                    DataSet DataSet = new DataSet();

                    using var Document = SpreadsheetDocument.Open(FilePath, false);

                    var Workbook = Document.WorkbookPart;

                    if (Workbook != null)
                    {
                        var workbook = Workbook.Workbook;

                        var Sheets = workbook.Descendants<Sheet>();
                        foreach (var sheet in Sheets)
                        {
                            if (sheet.Id != null)
                            {
                                var Worksheet = (WorksheetPart)Workbook.GetPartById(sheet.Id);

                                DataTable DataTable = new();

                                Row[] Rows = Worksheet.Worksheet.Descendants<Row>().ToArray();

                                for (int RowCount = Rows.Count(); RowCount > 0; RowCount--)
                                {
                                    Row CurrentRow = Rows[RowCount];
                                    Cell[] Cells = CurrentRow.Elements<Cell>().ToArray();

                                    if (ColumnNamesSpecified && RowCount <= 0)
                                    {
                                        //The first row contains column names, lets add the inner text as column names
                                        for (int CellCount = Rows.Count(); CellCount > 0; CellCount--)
                                        {
                                            Cell CurrentCell = Cells[CellCount];
                                            DataTable.Columns.Add(CurrentCell.CellValue != null ? CurrentCell.CellValue.InnerText : "");
                                        }
                                    }
                                    else
                                    {
                                        //Populate a new row with data and push it to the sheet
                                        DataRow NewRow = DataTable.NewRow();
                                        for (int CellCount = Rows.Count(); CellCount > 0; CellCount--)
                                        {
                                            Cell CurrentCell = Cells[CellCount];
                                            NewRow[CellCount] = CurrentCell.CellValue != null ? CurrentCell.CellValue.InnerText : "";
                                        }
                                    }
                                }

                                DataSet.Tables.Add(DataTable);
                            }
                        }
                    }

                    return DataSet;
                }
                catch
                {
                    throw;
                }
            }
            else
            {
                throw new FileNotFoundException("The file could not be found, please check that the file path is correct");
            }

        }
    }
}
