using Smartsheet.Api;
using Smartsheet.Api.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSheetAPI
{
        class Program
        {
            static Dictionary<string, long> columnMap = new Dictionary<string, long>();
            static long sheetId;
            public static PaginatedResult<Attachment> attachs;
            static void Main(string[] args)
            {
                SmartsheetClient client = new SmartsheetBuilder()
                    .SetAccessToken("aa63i6z8dyiol6brvs9gv92j9r")
                    .Build();

                PaginatedResult<Sheet> sheets = client.SheetResources.ListSheets(
                    null,
                    null,
                    null);

                Console.WriteLine("Found " + sheets.TotalCount);

                sheetId = (long)sheets.Data[0].Id;

                Sheet sheet = client.SheetResources.GetSheet(sheetId, null, null, null, null, null, null, null);

                foreach (var Row in sheet.Rows)
                {
                    attachs = client.SheetResources.RowResources.AttachmentResources.ListAttachments(sheetId, (long)Row.Id, null);
                }

     
                foreach (Column column in sheet.Columns)
                {
                    columnMap.Add(column.Title, (long)column.Id);
                }

                List<Row> rowsToUpdate = new List<Row>();

                foreach (Row row in sheet.Rows)
                {
                    Row rowToUpdate = evaluateRowAndBuildUpdate(row);
                    if (rowToUpdate != null)
                    {
                        rowsToUpdate.Add(rowToUpdate);
                    }
                }

                client.SheetResources.RowResources.UpdateRows(sheetId, rowsToUpdate);
                Console.ReadLine();
            }

            private static Row evaluateRowAndBuildUpdate(Row sourceRow)
            {
                Row rowToUpdate = null;

                Cell statusCell = getCellByColumnName(sourceRow, "Comments");
                if (statusCell.DisplayValue != null)
                {
                    var cellToUpdate = new Cell
                    {
                        ColumnId = columnMap["Comments"],
                        Value = "Prueba comentario 2"
                    };

                    var cellsToUpdate = new List<Cell>();
                    cellsToUpdate.Add(cellToUpdate);

                    rowToUpdate = new Row
                    {
                        Id = sourceRow.Id,
                        Cells = cellsToUpdate
                    };
                }

                return rowToUpdate;
            }

            private static Cell getCellByColumnName(Row row, string columnName)
            {
                var c = row.Cells.First(cell => cell.ColumnId == columnMap[columnName]);

                return c;
            }
        }
}
