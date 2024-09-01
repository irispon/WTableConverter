using Newtonsoft.Json.Linq;
using System;
using Microsoft.Extensions.Configuration;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using WonTableConverter.Excel2Data;
using WonTableConverter;
public class MainProgram
{

   public static Dictionary<string/*Sheet*/, Dictionary<string/*enumstring*/, int/*value*/>> enumDictionary = null;
    static void Main(string[] args)
    {

        enumDictionary = new Dictionary<string, Dictionary<string, int>>();
        string defaultExcelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelData\\");
        string defaultScriptOutPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertFolder\\");

        string appSettingsFilePath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");

        if (!File.Exists(appSettingsFilePath))
        {
            // 환경 설정 파일이 없으면 기본 경로 설정
            var defaultConfig = new
            {
                Paths = new
                {
                    ExcelFilePath = defaultExcelFilePath,
                    ScriptOutPath = defaultScriptOutPath
                }
            };

            // 기본 설정으로 appsettings.json 파일 생성
            var defaultConfigJson = System.Text.Json.JsonSerializer.Serialize(defaultConfig, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(appSettingsFilePath, defaultConfigJson);

            Console.WriteLine("appsettings.json 파일이 존재하지 않아 기본 설정으로 생성되었습니다.");
        }

        // 환경 설정 파일 읽기
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // 환경 설정에서 경로 읽기
        string excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), configuration["Paths:ExcelFilePath"]!);
        string scriptOutPath = Path.Combine(Directory.GetCurrentDirectory(), configuration["Paths:ScriptOutPath"]!); 
        string dataOutPath = scriptOutPath + "\\TableDatas\\";
        // 출력 폴더가 존재하지 않으면 생성합니다.
        if (!Directory.Exists(scriptOutPath))
        {
            Directory.CreateDirectory(scriptOutPath);
            Console.WriteLine($"{scriptOutPath} 폴더가 존재하지 않습니다.");
        }
        if (!Directory.Exists(dataOutPath))
        {
            Directory.CreateDirectory(dataOutPath);
            Console.WriteLine($"{dataOutPath} 폴더가 존재하지 않습니다.");
        }
        // 스크립트 클래스 생성
        var excelFiles = Directory.GetFiles(excelFilePath, "*.xlsx");
        var scriptCodeBuilder = new StringBuilder();

        if (Directory.Exists(excelFilePath))
        {
            if (excelFiles.Length < 1)
            {
                Console.WriteLine($"Excel file 이 존재하지 않습니다.");
                return;
            }
            List<string> _tmpFiles = new List<string>();
            try
            {


                foreach (var excelFile in excelFiles)
                {
                    if (excelFile.Contains("~$") == true)
                        continue;

                    string tempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileName(excelFile));
                    File.Copy(excelFile, tempFilePath, true);
                    _tmpFiles.Add(tempFilePath);
                }


                foreach (var excelFile in _tmpFiles)
                {
                    // 임시 파일 생성
                    ExcelToScriptConverter.GeneratedEnum(excelFile, enumDictionary);
                }


                foreach (var excelFile in _tmpFiles)
                {

                    // 스크립트 코드 생성 및 추가
                    string scriptCode = ExcelToScriptConverter.ConvertExcelToScript(excelFile);
                    scriptCodeBuilder.AppendLine(scriptCode);

                    // XML 데이터 생성
                    ExcelToXmlConverter.ConvertExcelToXml(excelFile, dataOutPath);

                }

                foreach (var excelFile in _tmpFiles)
                {
                    ExcelToXmlConverter.ConvertSubExcelToXml(excelFile, dataOutPath);
                }
            }
            finally
            {
                foreach(var excelFile in _tmpFiles)
                {
                    // 임시 파일 삭제
                    if (File.Exists(excelFile))
                    {
                        File.Delete(excelFile);
                    }
                }
   
            }
        }



        string scriptOutputFilePath = Path.Combine(scriptOutPath, "GeneratedScript.cs");
        File.WriteAllText(scriptOutputFilePath, scriptCodeBuilder.ToString());




        Console.WriteLine($"스크립트 및 JSON 데이터가 성공적으로 생성되었습니다: {scriptOutputFilePath}, {scriptOutPath}");
        Console.ReadLine();


    }
}
