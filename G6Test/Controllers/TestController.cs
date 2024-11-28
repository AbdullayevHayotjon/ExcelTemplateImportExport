using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace G6Test.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class TestController : ControllerBase
    {
        // Bu metod Excel faylini yaratib, uni foydalanuvchiga yuklab olish uchun taqdim etadi
        [HttpPost("ExportToExcel")]
        public async Task<IActionResult> ExportToExcel([FromBody] List<MyData> dataList)
        {
            // Excel faylini yaratish
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet 1");
                var currentRow = 1;

                // Sarlavhalar qo'shish (agar kerak bo'lsa)
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Name";
                worksheet.Cell(currentRow, 3).Value = "Role";

                // Data qo'shish
                foreach (var data in dataList)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = data.Id;
                    worksheet.Cell(currentRow, 2).Value = data.Name;
                    worksheet.Cell(currentRow, 3).Value = data.Role;
                }

                // Memory streamga yozish
                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    // Excel faylini yuklab olish uchun qaytarish
                    return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExportedData.xlsx");
                }
            }
        }

        [HttpPost("ImportFromExcel")]
        public async Task<IActionResult> ImportFromExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Fayl tanlanmagan.");

            try
            {
                // Excel faylini o'qish
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        var worksheet = workbook.Worksheet(1); // Birinchi varaqqa kirish
                        var rowCount = worksheet.RowsUsed().Count();

                        var dataList = new List<MyData>();

                        // Excel faylidan ma'lumotlarni o'qish
                        for (int row = 2; row <= rowCount; row++) // 2-qatordan boshlaymiz, chunki 1-qator sarlavhalar
                        {
                            var id = worksheet.Cell(row, 1).GetValue<int>(); // Id ni o'qish
                            var name = worksheet.Cell(row, 2).GetValue<string>(); // Name ni o'qish
                            var role = worksheet.Cell(row, 3).GetValue<string>(); // Role ni o'qish

                            dataList.Add(new MyData
                            {
                                Id = id,
                                Name = name,
                                Role = role
                            });
                        }

                        // List ni qaytarish
                        return Ok(dataList);
                    }
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Xatolik yuz berdi: {ex.Message}");
            }
        }

        [HttpPost("UploadTemplate")]
        public IActionResult UploadTemplate(IFormFile templateFile)
        {
            if (templateFile == null || templateFile.Length == 0)
                return BadRequest("Shablon fayli yuklanmadi!");

            try
            {
                var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");

                // Papka mavjud bo'lmasa, yaratish
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                var templatePath = Path.Combine(folderPath, "CurrentTemplate.xlsx");

                // Shablonni saqlash
                using (var stream = new FileStream(templatePath, FileMode.Create))
                {
                    templateFile.CopyTo(stream);
                }

                return Ok("Shablon muvaffaqqiyatli yuklandi.");
            }
            catch (Exception ex)
            {
                return BadRequest($"Xatolik yuz berdi: {ex.Message}");
            }
        }

        [HttpGet("DownloadTemplate")]
        public IActionResult DownloadTemplate()
        {
            try
            {
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "CurrentTemplate.xlsx");

                if (!System.IO.File.Exists(templatePath))
                    return NotFound("Shablon fayli mavjud emas!");

                var memory = new MemoryStream();
                using (var stream = new FileStream(templatePath, FileMode.Open))
                {
                    stream.CopyTo(memory);
                }

                memory.Position = 0;
                return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CurrentTemplate.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"Xatolik yuz berdi: {ex.Message}");
            }
        }

        [HttpPost("ValidateAndReadDynamicExcel")]
        public IActionResult ValidateAndReadDynamicExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("Fayl yuklanmadi!");

            try
            {
                // Shablon yo'lini olish
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Templates", "CurrentTemplate.xlsx");

                // Shablon fayl mavjudligini tekshirish
                if (!System.IO.File.Exists(templatePath))
                    return BadRequest("Shablon fayli mavjud emas!");

                // Shablondagi sarlavhalarni olish
                List<string> templateHeaders;
                using (var templateStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read))
                {
                    using (var templateWorkbook = new XLWorkbook(templateStream))
                    {
                        var templateWorksheet = templateWorkbook.Worksheet(1);
                        templateHeaders = templateWorksheet.Row(1).CellsUsed()
                                            .Select(c => c.Value.ToString().Trim())
                                            .ToList();
                    }
                }

                // Yuklangan faylni o'qish
                using (var userStream = new MemoryStream())
                {
                    file.CopyTo(userStream);
                    userStream.Seek(0, SeekOrigin.Begin);

                    using (var userWorkbook = new XLWorkbook(userStream))
                    {
                        var worksheet = userWorkbook.Worksheet(1);
                        var userHeaders = worksheet.Row(1).CellsUsed()
                                            .Select(c => c.Value.ToString().Trim())
                                            .ToList();

                        // Shablon bilan yuklangan faylni taqqoslash
                        if (!templateHeaders.SequenceEqual(userHeaders))
                        {
                            return BadRequest("Yuklangan fayl shablonga mos emas!");
                        }

                        // Fayldan ma’lumotlarni o‘qish
                        var dataList = new List<Dictionary<string, object>>();
                        foreach (var row in worksheet.RowsUsed().Skip(1)) // 1-qator sarlavha, keyingilar ma'lumotlar
                        {
                            var rowData = new Dictionary<string, object>();
                            for (int i = 0; i < userHeaders.Count; i++)
                            {
                                var header = userHeaders[i];
                                var cellValue = row.Cell(i + 1).Value.ToString(); // Hujayra qiymatini o‘qiymiz
                                rowData[header] = cellValue; // Sarlavha asosida qiymatni qo'shamiz
                            }
                            dataList.Add(rowData);
                        }

                        return Ok(dataList); // JSON formatda natijani qaytaramiz
                    }
                }
            }
            catch (Exception ex)
            {
                return BadRequest($"Xatolik yuz berdi: {ex.Message}");
            }
        }




    }
}
