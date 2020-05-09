using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace excelApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserController : ControllerBase
    {
        IUser _SUser;
        public UserController(IUser _SUser)
        {
            this._SUser = _SUser;
        }

        [HttpPost("dataimporttoexcelfile")]
        public IActionResult Import_Data()
        {
            List<User> users = _SUser.GetUsers();
            string fileName = "users.xls";
            WriteExcelFile(users,fileName);
            string currentDirectory = Directory.GetCurrentDirectory();
            byte[] arr = new byte[5000000];
            using (FileStream fs = System.IO.File.OpenRead($@"{currentDirectory}/{fileName}"))
            {
                arr = System.IO.File.ReadAllBytes($@"{currentDirectory}/{fileName}");
                fs.Close();
                System.IO.File.Delete($@"{currentDirectory}/{fileName}");
            }
            return File(arr, "application/vnd.ms-excel", fileName);
        }

        static void WriteExcelFile(List<User> users, string fileName)
        {
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(users), (typeof(DataTable)));
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
                workbookPart.Workbook.Save();
            }
        }
    }
}