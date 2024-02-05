using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json.Linq;
using System.Net.NetworkInformation;
using System.Threading;

string credPath = "C:\\Users\\RID - 4\\source\\repos\\GoogleSheetAPI\\CalculateApproval\\Credentials\\loyal-manifest-412818-b8aa1fa7e125.json";

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

string rangeOfNumberOfClasses = "engenharia_de_software!A2:H2";
SpreadsheetsResource.ValuesResource.GetRequest requestOfNumberOfClasses = service.Spreadsheets.Values.Get(spreadsheetId, rangeOfNumberOfClasses);
ValueRange responseOfNumberOfClasses = requestOfNumberOfClasses.Execute();
IList<IList<object>> valuesOfNumberOfClasses = responseOfNumberOfClasses.Values;

int numberOfClasses = 0;
if (valuesOfNumberOfClasses != null && valuesOfNumberOfClasses.Count > 0)
{
    foreach (var row in valuesOfNumberOfClasses)
    {
        foreach (var col in row)
        {
            if (col.ToString().Contains("Total de aulas no semestre:"))
            {
                string[] parts = col.ToString().Split(':');
                numberOfClasses = int.Parse(parts[1].Trim());
            }
        }
    }
}

Thread thread1 = new Thread(() => WriteApprovalStatusBasedOnAbsences(service, spreadsheetId, numberOfClasses));
Thread thread2 = new Thread(() => WriteApprovalStatusBasedOnGrades(service, spreadsheetId));
Thread thread3 = new Thread(() => checkFinalGrade(service, spreadsheetId));

thread1.Start();
thread2.Start();

thread1.Join();
thread2.Join();

thread3.Start();



static void WriteApprovalStatusBasedOnAbsences(SheetsService service, string spreadsheetId, int numberOfClasses)
{
    int rowIndex = 4;
    string rangeOfMissedClasses = "engenharia_de_software!C4:C";
    SpreadsheetsResource.ValuesResource.GetRequest requestOfMissedClasses = service.Spreadsheets.Values.Get(spreadsheetId, rangeOfMissedClasses);
    ValueRange responseOfMissedClasses = requestOfMissedClasses.Execute();
    IList<IList<object>> valuesOfMissedClasses = responseOfMissedClasses.Values;

    if (valuesOfMissedClasses != null && valuesOfMissedClasses.Count > 0)
    {
        foreach (var row in valuesOfMissedClasses)
        {
            int numberOfAbsences = 0;
            string status;

            foreach (var col in row)
            {
                if (int.TryParse(col.ToString(), out int missed))
                {
                    numberOfAbsences = missed;
                }
            }
            int numberOfMissedClassesAllowed = (numberOfClasses * 25) / 100;
            if (numberOfAbsences > numberOfMissedClassesAllowed)
            {
                status = "Reprovado por falta";
                string approvalStatusRange = "engenharia_de_software!G" + (rowIndex);
                var updateValues = new List<IList<object>> { new List<object> { status } };
                var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, approvalStatusRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var updateResponse = updateRequest.Execute();
                rowIndex++;
            }
            else
            {
                rowIndex++;
            }
        }
    }
}

static void WriteApprovalStatusBasedOnGrades(SheetsService service, string spreadsheetId)
{
    string range = "engenharia_de_software!D4:F";
    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
    ValueRange response = request.Execute();
    IList<IList<object>> values = response.Values;
    int rowIndex = 3;
    int student = 1;
    if (values != null && values.Count > 0)
    {
        foreach (var row in values)
        {
            double sum = 0;
            int count = 0;
            string status;
            foreach (var col in row)
            {
                if (double.TryParse(col.ToString(), out double grade))
                {
                    sum += grade;
                    count++;
                }
            }
            double average = Math.Ceiling(sum / count);
            Console.WriteLine("Média do aluno " + student++ + " : " + average);
            status = average >= 70.0 ? "Aprovado" : (average >= 50.0 ? "Exame Final" : "Reprovado por Nota");

            string approvalStatusRange = "engenharia_de_software!G" + (rowIndex + 1);
            SpreadsheetsResource.ValuesResource.GetRequest statusRequest =
            service.Spreadsheets.Values.Get(spreadsheetId, approvalStatusRange);
            ValueRange statusResponse = statusRequest.Execute();
            IList<IList<object>> statusValues = statusResponse.Values;
            
            if (statusValues == null)
            {
                var updateValues = new List<IList<object>> { new List<object> { status } };
                var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, approvalStatusRange);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var updateResponse = updateRequest.Execute();

            }
            else
            {
                status = statusValues[0][0].ToString();
            }
            if(status.CompareTo("Exame Final") == 0)
            {
                string finalGradeApproval = "engenharia_de_software!H" + (rowIndex + 1);
                var updateValues = new List<IList<object>> { new List<object> { average } };
                var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, finalGradeApproval);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var updateResponse = updateRequest.Execute();
            }
            rowIndex++;
        }
    }
}

static void checkFinalGrade(SheetsService service, string spreadsheetId)
{
    int rowIndex = 4;
    int i = 0;
    string approvalStatusRange = "engenharia_de_software!G4:G";
    SpreadsheetsResource.ValuesResource.GetRequest statusRequest = service.Spreadsheets.Values.Get(spreadsheetId, approvalStatusRange);
    ValueRange statusResponse = statusRequest.Execute();
    IList<IList<object>> statusValues = statusResponse.Values;
    if (statusValues != null)
    {
        foreach (var row in statusValues)
        {
            foreach (var col in row)
            {
                string status = statusValues[i++][0].ToString();
                if (status.CompareTo("Aprovado") == 0 || status.CompareTo("Reprovado por falta") == 0 || status.CompareTo("Reprovado por Nota") == 0)
                {
                    string finalGradeApproval = "engenharia_de_software!H" + rowIndex;
                    var updateValues = new List<IList<object>> { new List<object> { 0 } };
                    var updateRequest = service.Spreadsheets.Values.Update(new ValueRange { Values = updateValues }, spreadsheetId, finalGradeApproval);
                    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    var updateResponse = updateRequest.Execute();
                    rowIndex++;
                }
                else
                {
                    rowIndex++;
                }
            }
            
        }
        
    }
}










