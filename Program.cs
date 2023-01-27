using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;

namespace GoogleSheetsAndCsharp
{
    class Program
    {

        static readonly string[] Scopes ={SheetsService.Scope.Spreadsheets};
        static readonly string ApplicationName = "GenerateStats";
        static readonly string SpreadsheetId = "10j_UzKn7HaS-ysC7z64w9JzBK0f47ap4emMsWmrL2GM";
        static readonly string sheet = "Atributo";
        static SheetsService service;
        
        static void Main(string[] args){
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read)){
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer(){
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            ReadEntries();
        }
        static void ReadEntries(){
            var range = $"{sheet}!A1:J141";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0){
                foreach (var row in values){
                    Console.WriteLine("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8}", row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9]);
                }
            }
            else{
                Console.WriteLine("No data found");
            }
        }
    }
}