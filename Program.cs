using System;
using System.IO;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;

namespace GoogleSheetsAndCsharp
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "GenerateStats";
        static readonly string SpreadsheetId = "10j_UzKn7HaS-ysC7z64w9JzBK0f47ap4emMsWmrL2GM";
        static readonly string sheet = "Atributo";
        static SheetsService service;

        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }
            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Deserialize JSON into an object
            var jsonString = File.ReadAllText("data.json");
            var data = JsonConvert.DeserializeObject<Data>(jsonString);
            Console.WriteLine("data:");
            Console.WriteLine(data);
            // Read entries that match the specified Speed and GuardianClass
            var range = $"{sheet}!A1:J141";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request.Execute();
            var values = response.Values;
            if (values != null && values.Count > 0)
            {
                var matchingRows = values.Where(row => row[1].ToString() == data.properties.GuardianClass && row[9].ToString() == data.properties.Stats.Speed.ToString());
                if (matchingRows.Any())
                {
                    // Choose a random row that matches the criteria
                    var randomRow = matchingRows.ElementAt(new Random().Next(matchingRows.Count()));

                    // Create a new JSON object with the selected row's values
                    var resultData = new ResultData
                    {
                        HP = int.Parse(randomRow[2].ToString()),
                        Strength = int.Parse(randomRow[3].ToString()),
                        Resolve = int.Parse(randomRow[4].ToString()),
                        Armor = int.Parse(randomRow[5].ToString()),
                        Mantle = int.Parse(randomRow[6].ToString()),
                        MaxSP = int.Parse(randomRow[7].ToString()),
                        TurnSP = int.Parse(randomRow[8].ToString()),
                        Speed = int.Parse(randomRow[9].ToString())
                    };
                    // Create a new JSON object with the selected row's values, and additional information from data.json
                    
                    data.properties.Stats.HP = resultData.HP;
                    data.properties.Stats.Strength = resultData.Strength;
                    data.properties.Stats.Resolve = resultData.Resolve;
                    data.properties.Stats.Armor = resultData.Armor;
                    data.properties.Stats.Mantle = resultData.Mantle;
                    data.properties.Stats.MaxSP = resultData.MaxSP;
                    data.properties.Stats.TurnSP = resultData.Speed;
                    // Serialize the object to a JSON string
                    var resultJson = JsonConvert.SerializeObject(resultData);
                    var rookieResultJson = JsonConvert.SerializeObject(data);
                    Console.WriteLine("data:");
                    Console.WriteLine(data);
                    // Write the JSON string to a file
                    File.WriteAllText("result.json", resultJson);
                    File.WriteAllText("rookie_result.json", rookieResultJson);

                    // Print the JSON string to the console
                    Console.WriteLine("data:");
                    Console.WriteLine(data);
                    Console.WriteLine("Result:");
                    Console.WriteLine(resultJson);
                    Console.WriteLine("Rookie Result:");
                    Console.WriteLine(rookieResultJson);
                }
                else
                {
                    Console.WriteLine("No matching rows found.");
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.ReadLine();
        }


    }


    class Data
    {
        public string id { get; set; }
        public string name {get;set;}
        public string description {get;set;}
        public string image {get;set;}
        public int nFT_Type{get;set;}
        public Properties properties { get; set; }
    }

    class Properties
    {
        public string PlushieDomain {get; set;}
        public string GuardianClass { get; set; }
        public Stats Stats { get; set; }
        public int Trait1{get;set;}
        public int Trait2{get;set;}
        public int Trait3{get;set;}
        public int Trait4{get;set;}
    }

    class Stats
    {
        public int HP {get;set;}
        public int Strength{get;set;}
        public int Resolve {get;set;}
        public int Armor{get;set;}
        public int Mantle{get;set;}
        public int MaxSP{get;set;}
        public int TurnSP{get;set;}
        public int Speed { get; set; }
    }

    class ResultData
    {
        public int HP { get; set; }
        public int Strength { get; set; }
        public int Resolve { get; set; }
        public int Armor { get; set; }
        public int Mantle { get; set; }
        public int MaxSP { get; set; }
        public int TurnSP { get; set; }
        public int Speed { get; set; }
    }

    
}