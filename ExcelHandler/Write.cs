using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHandler
{
    public class Write
    {
        public static void DatasetToExcel(DataSet DataSet, string SavePath, bool IncludeHeaderRow = false)
        {
            try
            {
                using (var workbook = SpreadsheetDocument.Create(SavePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = workbook.AddWorkbookPart();

                    workbook.WorkbookPart.Workbook = new();

                    workbook.WorkbookPart.Workbook.Sheets = new();

                    foreach (DataTable Table in DataSet.Tables)
                    {

                        var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                        Sheet Sheet = new();
                        sheetPart.Worksheet = new(Sheet);

                        Sheets Sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                        string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                        uint sheetId = 1;
                        if (Sheets.Elements<Sheet>().Count() > 0)
                        {
                            sheetId =
                                Sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                        }

                        Sheet sheet = new() { Id = relationshipId, SheetId = sheetId, Name = Table.TableName };
                        Sheets.Append(sheet);

                        List<string> Columns = new List<string>();

                        if (IncludeHeaderRow)
                        {
                            Row headerRow = new();

                            foreach (DataColumn column in Table.Columns)
                            {
                                Columns.Add(column.ColumnName);

                                Cell cell = new();
                                cell.DataType = CellValues.String;
                                cell.CellValue = new CellValue(column.ColumnName);
                                headerRow.AppendChild(cell);
                            }

                            Sheet.AppendChild(headerRow);
                        }


                        foreach (DataRow dsrow in Table.Rows)
                        {
                            Row newRow = new();
                            foreach (string Column in Columns)
                            {
                                Cell cell = new();
                                cell.DataType = CellValues.String;
                                cell.CellValue = new(dsrow[Column].ToString());
                                newRow.AppendChild(cell);
                            }

                            Sheet.AppendChild(newRow);
                        }

                    }
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
