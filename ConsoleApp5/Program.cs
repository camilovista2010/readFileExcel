using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;

namespace ConsoleApp5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            ProcessFile processFile = new ProcessFile();
            processFile.getFileAsync();

            Console.ReadLine(); 
           
        }


    }


    public class ProcessFile 
    {
        public async void getFileAsync()
        {
            HttpClient _client = new HttpClient(); 

            try
            {
                var filePath = "https://businttidiag.blob.core.windows.net/posint/TestingExcel.xlsx";
                var nameColumn = "B";
                var nameSheet = "Hoja1";

                HttpResponseMessage task = await _client.GetAsync(filePath);
                Stream stream = await task.Content.ReadAsStreamAsync();
                 

                using (var document = SpreadsheetDocument.Open(stream, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var workbook = workbookPart.Workbook;

                    var sheetsResult = workbook.Descendants<Sheet>().Where(y => y.Name == nameSheet).FirstOrDefault();

                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetsResult.Id);
                    var sharedStringPart = workbookPart.SharedStringTablePart;
                    var values = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

                    var cells = worksheetPart.Worksheet.Descendants<Cell>();
                    foreach (var cell in cells)
                    {
                        if (cell.CellReference.Value.StartsWith(nameColumn))
                        {
                            // The cells contains a string input that is not a formula
                            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                            {
                                var index = int.Parse(cell.CellValue.Text);
                                var value = values[index].InnerText;
                                Console.WriteLine(value);
                            }
                            else
                            {
                                Console.WriteLine(cell.CellValue.Text);
                            }

                            if (cell.CellFormula != null)
                            {
                                Console.WriteLine(cell.CellFormula.Text);
                            }
                        }

                    }

                    var countCellValues = cells.Where(x => x.CellReference.Value.StartsWith(nameColumn)).Count();

                    Console.WriteLine(countCellValues);
                }
            }
            catch
            {

            }
        }
    
    }
}
