using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

//points to the credential provided by Google to access the spreadsheet (Change to your local credential)
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

//Id of ur spreadSheet (Change for the sheet's id that you want to modify)
string spreadsheetId = "1-iz4w7WxPP5Dj6I-UjL5t1Y9WDrXqgg_g7B2csJVRbo";
int numberOfClasses = 0;

//Defined the range that i want to acess
string rangeOfNumberOfClasses = "engenharia_de_software!A2:H2";

//Acessing a specify range to get the maximum number of classes
SpreadsheetsResource.ValuesResource.GetRequest requestOfNumberOfClasses =
service.Spreadsheets.Values.Get(spreadsheetId, rangeOfNumberOfClasses);
ValueRange responseOfNumberOfClasses = requestOfNumberOfClasses.Execute();
IList<IList<object>> valuesOfNumberOfClasses = responseOfNumberOfClasses.Values;

//Scrolling through the spreadSheet 
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

//Function that Update the SpreadSheet
UpdateSheet(service, spreadsheetId, numberOfClasses);

static void UpdateSheet(SheetsService service, string spreadsheetId, int numberOfClasses)
{
    int students = 1;
    //Setting the range to acess missed classes and all three grades
    string range = "engenharia_de_software!C4:F";
    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, range);
    ValueRange response = request.Execute();
    IList<IList<object>> values = response.Values;

    if (values != null && values.Count > 0)
    {
        //Setting the Lists of Status and Final Grade to update on the spreadSheet
        List<IList<object>> updateStatus = new List<IList<object>>();
        List<IList<object>> updateFinalGrade = new List<IList<object>>();

        try
        {
            //Scrolling through the rows
            foreach (var row in values)
            {

                //Getting the number of Abscences
                int numberOfAbsences = int.Parse(row[0].ToString());

                //Getting the Grades
                double p1 = double.Parse(row[1].ToString());
                double p2 = double.Parse(row[2].ToString());
                double p3 = double.Parse(row[3].ToString());

                //Calculating the number of Missed Classes Allowed
                int numberOfMissedClassesAllowed = (numberOfClasses * 25) / 100;

                double sum = p1 + p2 + p3;
                //calculating the average  
                double average = Math.Ceiling(sum / 3);
                //Printing on Console the average of each student
                Console.WriteLine("Student " + students + " average: " + average);
                students++;
                //Setting the status through the average and the absences
                string status = average >= 70.0 ? "Aprovado" : (average >= 50.0 ? "Exame Final" : "Reprovado por Nota");
                status = numberOfAbsences > numberOfMissedClassesAllowed ? "Reprovado por falta" : status;
                double finalGrade = status == "Exame Final" ? 100 - average : 0;

                //Adding the values of Status and Final Grades to the Lists 
                updateStatus.Add(new List<object> { status });
                updateFinalGrade.Add(new List<object> { finalGrade });
            }
        }catch(FormatException ex)
        {
            Console.WriteLine("Verifique se todas as células de faltas e de notas são de fato números");
        }

        //Using class Batch to update the spreadsheet at once
        BatchUpdateValuesRequest batchRequest = new BatchUpdateValuesRequest();
        batchRequest.ValueInputOption = "RAW";
        //Setting a new Object as a parameter (Range, Values)
        batchRequest.Data = new List<ValueRange> 
    {
        new ValueRange { Range = "engenharia_de_software!G4:G", Values = updateStatus.Take(values.Count).ToList() },
        new ValueRange { Range = "engenharia_de_software!H4:H", Values = updateFinalGrade.Take(values.Count).ToList() }
    };
        //Running Batch to Update the SpreadSheet
        service.Spreadsheets.Values.BatchUpdate(batchRequest, spreadsheetId).Execute();
    }
}








