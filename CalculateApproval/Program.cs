using CalculateApproval;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.IO;

string credPath = "/Users/caiotbraga/Desktop/GOOGLE/GoogleSheetAPI/CalculateApproval/Credentials/loyal-manifest-412818-b8aa1fa7e125.json";

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
int rowIndex = 3;

string rangeOfMissedClasses= "engenharia_de_software!C:C";
SpreadsheetsResource.ValuesResource.GetRequest requestOfMissedClasses = service.Spreadsheets.Values.Get(spreadsheetId, rangeOfMissedClasses);
ValueRange responseOfMissedClasses = requestOfMissedClasses.Execute();
IList<IList<object>> valuesOfMissedClasses = responseOfMissedClasses.Values;

if(valuesOfMissedClasses != null && valuesOfMissedClasses.Count > 0)
{
    foreach (var row in valuesOfMissedClasses)
    {
        int numberOfClasses = 0;
        int numberOfAbsences = 0;
        string status;

        foreach (var col in row)
        {
            if (col.ToString().Contains("Total de aulas no semestre:"))
            {
                string[] parts = col.ToString().Split(':');
                numberOfClasses = int.Parse(parts[1].Trim());
            }
            else if (int.TryParse(col.ToString(), out int missed))
            {
                numberOfAbsences = missed;
            }
        }
        
        int numberOfMissedClassesAllowed = (numberOfClasses * 25) / 100;

        if (numberOfAbsences > numberOfMissedClassesAllowed)
        {
            status = "Reprovado por falta";
            string approvalStatusRange = "engenharia_de_software!G" + (rowIndex + 1);
            var updateValues = new List<IList<object>> { new List<object> { status } };
            var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, approvalStatusRange);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var updateResponse = updateRequest.Execute();
        }
    }
}

string range = "engenharia_de_software!D4:F";
SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
ValueRange response = request.Execute();
IList<IList<object>> values = response.Values;

if (values != null && values.Count > 0)
{
    foreach (var row in values)
    {
        double sum = 0;
        int count = 0;
        double average = sum / count;
        string status;
        foreach (var col in row)
        {
            if (double.TryParse(col.ToString(), out double grade))
            {
                sum += grade;
                count++;
            }
        }
        status = average >= 70.0 ? "Aprovado" : (average >= 50.0 ? "Exame Final" : "Reprovado por Nota");
        string approvalStatusRange = "engenharia_de_software!G" + (rowIndex + 1);
        SpreadsheetsResource.ValuesResource.GetRequest statusRequest =
        service.Spreadsheets.Values.Get(spreadsheetId, approvalStatusRange);
        ValueRange statusResponse = statusRequest.Execute();
        IList<IList<object>> statusValues = statusResponse.Values;
        if(statusValues == null)
        {
            var updateValues = new List<IList<object>> { new List<object> { status } };
            var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, approvalStatusRange);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var updateResponse = updateRequest.Execute();
        }
        rowIndex++;
    }
}
else
{
    Console.WriteLine("Nenhum dado encontrado.");
    
}

