using System.Text;
using ExcelReader;
using Newtonsoft.Json;
using OfficeOpenXml;

var userCarInfos = new List<UserCarInfo>();

ExcelPackage.LicenseContext = LicenseContext.Commercial;
var file = new FileInfo("C:\\Users\\Goga\\RiderProjects\\ExcelReader\\ExcelReader\\test.xlsx");

using var package = new ExcelPackage(file);
var worksheet = package.Workbook.Worksheets[0];
var rowCount = worksheet.Dimension.Rows;
var colCount = worksheet.Dimension.Columns;
for (var row = 1; row <= rowCount; row++)
{
    var userCarInfo = new UserCarInfo();
    for (var col = 1; col <= colCount; col++)
    {
        var data = worksheet.Cells[row, col].Value;
        if (col % 2 == 1)
            userCarInfo.CarName = data.ToString()!.Trim();
        else
            userCarInfo.CardId = data.ToString()!.Trim();
        
        userCarInfos.Add(userCarInfo);
    }
}

if (userCarInfos.Count != 0)
{
    using var client = new HttpClient();

    var content = new StringContent(
        JsonConvert.SerializeObject(userCarInfos),
        Encoding.UTF8,
        "application/json"
    );

    var response = await client.PostAsync("https://localhost:7157/api/excel", content);
    var responseContent = response.Content.ReadAsStringAsync().Result;
    Console.WriteLine(responseContent);
}
    
