using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ApiExplorer;
using ExcelDataReader;
using Service1.Models;
using Microsoft.Extensions.Options;

namespace Service1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DataOfExcelController : ControllerBase
    {
        private readonly FileSettings _fileSettings;
        string fullPath ;
        private readonly AppDbContext _context;
        public DataOfExcelController(IOptions<FileSettings> fileSettings,
                                     AppDbContext context)
        {
            _fileSettings = fileSettings.Value;
            _context = context;
        }
        [HttpPost]
        public IActionResult UploadSheet(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return NoContent();
            }

            var extension = Path.GetExtension(file.FileName);

            var fileName = $"{Guid.NewGuid()}{extension}";

            string? folderName = _fileSettings.Path + "/Data";

            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }

            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            fullPath = Path.Combine(pathToSave, fileName);
            var dbPath = Path.Combine(folderName, fileName);

            using (var fileStream = new FileStream(fullPath, FileMode.Create))
            {
                file.CopyTo(fileStream);
            }
            _context.Data.RemoveRange(_context.Data.Where(x=>x.Id >0 ));
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = System.IO.File.Open(fullPath, FileMode.Open, FileAccess.Read))
            {

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var confg = new ExcelDataSetConfiguration();
                    var dataSet = reader.AsDataSet(confg);
                    var count = dataSet.Tables[0].Rows.Count;
                    for (int i = 1; i < count; i++)
                    {
                        var row = new DataSheet
                        {
                            Id = Convert.ToInt32(dataSet.Tables[0].Rows[i][0]),
                            Name = dataSet.Tables[0].Rows[i][1].ToString(),
                            Description = dataSet.Tables[0].Rows[i][2].ToString(),
                            Location = dataSet.Tables[0].Rows[i][3].ToString(),
                            Price =Convert.ToDouble(dataSet.Tables[0].Rows[i][4]),
                            Color = dataSet.Tables[0].Rows[i][5].ToString()
                        };

                        _context.Data.Add(row);
                        _context.SaveChanges();
                    }

                }
                return Ok(dbPath);

            }
        }
        [HttpGet("/{Id}")]
        public IActionResult GetDataFromSheet([FromRoute] int Id)
        {
           var record =  _context.Data.FirstOrDefault(x => x.Id == Id);
           if(record!=null) return Ok(record);
            return BadRequest();
        }
           
    }
}
