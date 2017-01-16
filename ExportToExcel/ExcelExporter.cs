using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Web;


namespace ExportToExcel
{
    public static class ExcelExporter
    {
        public static void ExportDSToExcel<T>(this HttpResponseBase Response, IList<T> items, string destination)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                // using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                using (var workbook = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook, true))
                {
                    var workbookPart = workbook.AddWorkbookPart();
                    workbook.WorkbookPart.Workbook = new Workbook();
                    workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                    uint sheetId = 1;

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Any())
                    {
                        sheetId =
                            sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = typeof(T).Name };
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    Dictionary<string, PropertyInfo> columns = new Dictionary<string, PropertyInfo>();
                    var properties = typeof(T).GetProperties();
                    foreach (var prop in properties)
                    {
                        columns.Add(prop.Name, prop);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(prop.Name);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (var item in items)
                    {
                        Row newRow = new Row();
                        foreach (var col in columns.Keys)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(columns[col].GetValue(item).ToString());
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
                stream.Flush();
                stream.Position = 0;

                Response.ClearContent();
                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache);
                Response.AddHeader("content-disposition", "attachment; filename=" + destination);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }
    }
}
