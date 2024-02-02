using CalculateApproval;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;

string credPath = "C:\\Users\\RID - 4\\Source\\Repos\\GoogleSheetAPI\\CalculateApproval\\Credentials\\loyal-manifest-412818-b8aa1fa7e125.json";

string[] scopes = { SheetsService.Scope.Spreadsheets};

GoogleCredential credential;
using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
{
    credential = GoogleCredential.FromStream(stream)
        .CreateScoped(scopes);
}

var service = new SheetsService(new BaseClientService.Initializer()
{
    HttpClientInitializer = credential,
    ApplicationName = "CalculateApproval",
});

string spreadsheetId = "1-iz4w7WxPP5Dj6I-UjL5t1Y9WDrXqgg_g7B2csJVRbo";

string range = "engenharia_de_software!D4:F4";

SpreadsheetsResource.ValuesResource.GetRequest request =
        service.Spreadsheets.Values.Get(spreadsheetId, range);

ValueRange response = request.Execute();
IList<IList<object>> values = response.Values;

if (values != null && values.Count > 0)
{
    int rowIndex = 3;
    foreach (var row in values)
    {
        double sum = 0;
        int count = 0;
        foreach (var col in row)
        {
            if (double.TryParse(col.ToString(), out double grade))
            {
                sum += grade;
                count++;
            }
        }

        double average = sum / count;

        string studentAverageRange = "engenharia_de_software!H" + (rowIndex + 1); 
        var updateValues = new List<IList<object>> { new List<object> { average } };
        var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, studentAverageRange);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        var updateResponse = updateRequest.Execute();

        string status = average >= 70.0 ? "Aprovado" : (average >= 50.0 ? "Exame Final" : "Reprovado por Nota");

        string approvalStatusRange = "engenharia_de_software!G" + (rowIndex + 1); 
        updateValues = new List<IList<object>> { new List<object> { status } };
        updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, approvalStatusRange);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        updateResponse = updateRequest.Execute();

        rowIndex++;
    }
}
else
{
    Console.WriteLine("Nenhum dado encontrado.");
}

