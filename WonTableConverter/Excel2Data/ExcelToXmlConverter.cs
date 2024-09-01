using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace WonTableConverter.Excel2Data
{
    class ExcelToXmlConverter
    {
        public static void ConvertExcelToXml(string filePath, string outputDirectory)
        {
            // 엑셀 파일을 엽니다.
            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // '_sub'가 워크시트 이름에 포함되어 있으면 건너뜀
                    if (worksheet.Name.Contains("_sub", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;  // 이 워크시트는 ConvertSubExcelToXml에서 처리하도록 건너뜁니다.
                    }

                    // XML 루트 요소 생성
                    var root = new XElement("Root");

                    // 첫 번째 행: 변수명
                    var variableNames = worksheet.Row(1).CellsUsed().Select(c => c.GetValue<string>()).ToList();
                    var variableValues = worksheet.Row(4).CellsUsed().Select(c => c.GetValue<string>()).ToList();

                    if (variableNames.Count != variableValues.Count)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"[W] 해당 Sheet {worksheet.Name}의 1행과 3행의 숫자가 일치하지 않음");
                        Console.ResetColor(); // 출력 색상을 기본값으로 돌려놓습니다.          
                    }

                    // 두 번째 행: ignore 및 thirdkey 여부 확인
                    List<string>? ignoreFlags = new List<string>();
                    for (int i = 0; i < variableNames.Count; i++)
                    {
                        var cellValue = worksheet.Row(2).Cell(i + 1).GetValue<string>(); // 열을 1부터 시작하므로 i + 1
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            ignoreFlags.Add("None");
                        }
                        else
                        {
                            ignoreFlags.Add(cellValue.ToLower());
                        }
                    }

                    // 4행부터 데이터를 읽어들입니다.
                    for (int rowIdx = 5; rowIdx <= worksheet.LastRowUsed().RowNumber(); rowIdx++)
                    {
                        var rowElement = new XElement("Row");
                        var row = worksheet.Row(rowIdx);
                        for (int i = 0; i < variableNames.Count; i++)
                        {
                            // "ignore"인 열은 무시
                            if (ignoreFlags[i].Equals("ignore", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            string variableName = variableNames[i];
                            string variableValue = variableValues[i];
                            var cellValue = row.Cell(i + 1).Value;
                            string? cellText = row.Cell(i + 1).GetValue<string>();
                            if (string.IsNullOrEmpty(cellText) == true)
                            {
                                cellText = "";
                            }
                            if (variableValue.Contains("enum", StringComparison.OrdinalIgnoreCase))
                            {
                                string? cellName = worksheet.Row(3).Cell(i + 1).GetValue<string>();
                                if (string.IsNullOrEmpty(cellName) == true)
                                {
                                    cellName = worksheet.Name;
                                }

                                if (MainProgram.enumDictionary.ContainsKey(cellName) && MainProgram.enumDictionary[cellName].ContainsKey(cellText))
                                {
                                    cellValue = MainProgram.enumDictionary[cellName][cellText];
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine($"[W] 해당 Sheet {cellName}과 매치되지 않습니다. {cellText}가 이상함");
                                    Console.ResetColor(); // 출력 색상을 기본값으로 돌려놓습니다.                                }

                                }

                            }


                            // XML 요소로 변환
                            var cellElement = new XElement(variableName, cellValue);
                            rowElement.Add(cellElement);
                        }

                        root.Add(rowElement);
                    }

                    // 시트 이름을 파일명으로 사용하여 XML 파일 저장
                    string sheetFileName = worksheet.Name + ".xml";
                    string outputPath = Path.Combine(outputDirectory, sheetFileName);
                    root.Save(outputPath);

                    Console.WriteLine($"XML 파일이 생성되었습니다: {outputPath}");
                }
            }
        }
        public static void ConvertSubExcelToXml(string filePath, string outputDirectory)
        {
            // 엑셀 파일을 엽니다.
            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // '_sub'가 포함된 워크시트만 처리
                    if (!worksheet.Name.Contains("_sub", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // 기존에 생성된 XML 파일을 불러옴
                    string originalFileName = worksheet.Name.Replace("_sub", "") + ".xml";
                    string originalFilePath = Path.Combine(outputDirectory, originalFileName);

                    XElement root;
                    if (File.Exists(originalFilePath))
                    {
                        root = XElement.Load(originalFilePath);  // 기존 XML 파일을 불러옴
                    }
                    else
                    {
                        root = new XElement("Root");  // XML 파일이 없으면 새로 생성
                    }

                    // 첫 번째 행: 변수명
                    var variableNames = worksheet.Row(1).CellsUsed().Select(c => c.GetValue<string>()).ToList();
                    var variableValues = worksheet.Row(4).CellsUsed().Select(c => c.GetValue<string>()).ToList();

                    if (variableNames.Count != variableValues.Count)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"[W] 해당 Sheet {worksheet.Name}의 1행과 4행의 열 수가 일치하지 않습니다.");
                        Console.ResetColor(); // 출력 색상을 기본값으로 돌려놓습니다.
                    }

                    // 두 번째 행: ignore 및 thirdkey 여부 확인
                    List<string> ignoreFlags = new List<string>();
                    for (int i = 0; i < variableNames.Count; i++)
                    {
                        var cellValue = worksheet.Row(2).Cell(i + 1).GetValue<string>(); // 열을 1부터 시작하므로 i + 1
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            ignoreFlags.Add("None");
                        }
                        else
                        {
                            ignoreFlags.Add(cellValue.ToLower());
                        }
                    }

                    // 5행부터 데이터를 읽어들입니다.
                    for (int rowIdx = 5; rowIdx <= worksheet.LastRowUsed().RowNumber(); rowIdx++)
                    {
                        var rowElement = new XElement("Row");
                        var row = worksheet.Row(rowIdx);
                        for (int i = 0; i < variableNames.Count; i++)
                        {
                            // "ignore"인 열은 무시
                            if (ignoreFlags[i].Equals("ignore", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            string variableName = variableNames[i];
                            string variableValue = variableValues[i];
                            var cellValue = row.Cell(i + 1).Value;
                            string cellText = row.Cell(i + 1).GetValue<string>();

                            if (string.IsNullOrEmpty(cellText))
                            {
                                cellText = "";
                            }

                            if (variableValue.Contains("enum", StringComparison.OrdinalIgnoreCase))
                            {
                                string cellName = worksheet.Row(3).Cell(i + 1).GetValue<string>();
                                if (string.IsNullOrEmpty(cellName))
                                {
                                    cellName = worksheet.Name;
                                }

                                if (MainProgram.enumDictionary.ContainsKey(cellName) && MainProgram.enumDictionary[cellName].ContainsKey(cellText))
                                {
                                    cellValue = MainProgram.enumDictionary[cellName][cellText];
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine($"[W] 해당 Sheet {cellName}과 매치되지 않습니다. {cellText}가 이상함");
                                    Console.ResetColor(); // 출력 색상을 기본값으로 돌려놓습니다.
                                }
                            }

                            // XML 요소로 변환
                            var cellElement = new XElement(variableName, cellValue);
                            rowElement.Add(cellElement);
                        }

                        root.Add(rowElement);
                    }

                    // 기존 파일에 데이터를 추가하고 저장
                    root.Save(originalFilePath);
                    Console.WriteLine($"XML 파일에 데이터를 추가했습니다: {originalFilePath}");
                }
            }
        }



    }
}
