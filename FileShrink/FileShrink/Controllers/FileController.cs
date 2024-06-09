using System.IO;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace FileShrink.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private readonly string _fileName = "output.xlsx";
        private readonly string _wwwrootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");

        [HttpGet("generate")]
        public IActionResult GenerateExcel()
        {
            var data = new dynamic[,]
            {
                { "Header1", "Header2", "Header3" },
                { "Row1Col1", "Row1Col2", "Row1Col3" },
                { "Row2Col1", "Row2Col2", "Row2Col3" },
                { "Row3Col1", "Row3Col2", "Row3Col3" }
            };

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < data.GetLength(0); i++)
                {
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        worksheet.Cell(i + 1, j + 1).Value = data[i, j];
                    }
                }

                // Ensure the directory exists
                Directory.CreateDirectory(_wwwrootPath);
                string filePath = Path.Combine(_wwwrootPath, _fileName);

                // Save the workbook
                workbook.SaveAs(filePath);

                return Ok(new { FilePath = $"/{_fileName}" });
            }
        }

        [HttpDelete("delete")]
        public IActionResult DeleteExcel()
        {
            // Define the path to the Excel file
            string filePath = Path.Combine(_wwwrootPath,_fileName);
            

            if (System.IO.File.Exists(filePath))
            {
                // Delete the file
                System.IO.File.Delete(filePath);
                return Ok(new { Message = "File deleted successfully." });
            }
            else
            {
                return NotFound(new { Message = "File not found." });
            }
        }
    }
}
