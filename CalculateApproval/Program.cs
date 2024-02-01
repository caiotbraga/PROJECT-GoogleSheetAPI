using CalculateApproval;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;

string credPath = "C:\\Users\\RID - 4\\Desktop\\GoogleSheet\\GoogleSheetAPI\\CalculateApproval\\Credentials\\loyal-manifest-412818-b8aa1fa7e125.json"; 

string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };

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

string range = "engenharia_de_software!D4:E4"; 

SpreadsheetsResource.ValuesResource.GetRequest request =
        service.Spreadsheets.Values.Get(spreadsheetId, range);

ValueRange response = request.Execute();
IList<IList<object>> values = response.Values;

if (values != null && values.Count > 0)
{
    Console.WriteLine("Valores lidos da planilha:");
    foreach (var row in values)
    {
        foreach (var col in row)
        {
            Console.Write("{0} ", col);
        }
        Console.WriteLine();
    }
}
else
{
    Console.WriteLine("Nenhum dado encontrado.");
}

