using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text;

namespace WonTableConverter
{
    class ExcelToScriptConverter
    {
        public static string ConvertExcelToScript(string filePath)
        {
            var scriptCodeBuilder = new StringBuilder();

            // 엑셀 파일을 엽니다.
            using (var workbook = new XLWorkbook(filePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Contains("_sub", StringComparison.OrdinalIgnoreCase))
                    {
                        //[W] 서브 테이블인 경우는 스크립트를 생성하지 않음.
                        continue;
                    }
                    // 클래스 이름으로 시트 이름 사용
                    string className = worksheet.Name;
                    scriptCodeBuilder.AppendLine($"public class {className}");
                    scriptCodeBuilder.AppendLine("{");

                    // 첫 번째 행: 변수명
                    var variableNames = worksheet.Row(1).CellsUsed().Select(c => c.GetValue<string>()).ToList();

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
                    //


                    // 세 번째 행: 변수 타입
                    var variableTypes = worksheet.Row(4).CellsUsed().Select(c => c.GetValue<string>()).ToList();

                    // 속성 정의
                    for (int i = 0; i < variableNames.Count; i++)
                    {
                        // 빈 셀의 경우 기본값으로 처리
                        if (i >= ignoreFlags.Count || string.IsNullOrEmpty(ignoreFlags[i]))
                        {
                            ignoreFlags.Add("");  // 기본 빈 값을 추가
                        }

                        // "ignore"인 열은 무시
                        if (ignoreFlags[i].Equals("ignore", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string variableName = variableNames[i];
                        string variableType = variableTypes[i];

                        // 변수 타입에 따라 기본값을 설정합니다.
                        string defaultValue = variableType switch
                        {
                            "string" => "\"\"",
                            "int" => "0",
                            "enum" => "0",
                            "float" => "0.0f",
                            "byte" => "0",
                            "bool" => "false",
                            _ => "\"\"" // 기본값으로 string으로 처리
                        };
                        if (variableType == "enum")
                        {
                            variableType = "int";
                        }
                        // 변수 선언 코드 작성
                        scriptCodeBuilder.AppendLine($"    public {variableType} {variableName} {{ get; set; }} = {defaultValue};");
                    }

                    // thirdkey 존재 여부에 따라 반환 타입 결정
                    bool hasThirdKey = ignoreFlags.Contains("thirdkey", StringComparer.OrdinalIgnoreCase);
                    string returnType = hasThirdKey ? $"Dictionary<int, List<{className}>>" : $"Dictionary<int, {className}>";

                    // CreateData 메서드 생성
                    scriptCodeBuilder.AppendLine();
                    scriptCodeBuilder.AppendLine($"    public static {returnType} CreateData(string xmlPath)");
                    scriptCodeBuilder.AppendLine("    {");
                    scriptCodeBuilder.AppendLine(hasThirdKey
                        ? $"        var dictionary = new Dictionary<int, List<{className}>>();"
                        : $"        var dictionary = new Dictionary<int, {className}>();");
                    scriptCodeBuilder.AppendLine("        var document = System.Xml.Linq.XDocument.Load(xmlPath);");
                    scriptCodeBuilder.AppendLine("        foreach (var rowElement in document.Descendants(\"Row\"))");
                    scriptCodeBuilder.AppendLine("        {");
                    scriptCodeBuilder.AppendLine($"            var instance = new {className}();");

                    // 속성에 XML 값 할당
                    for (int i = 0; i < variableNames.Count; i++)
                    {
                        // "ignore"인 열은 무시
                        if (ignoreFlags[i].Equals("ignore", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string variableName = variableNames[i];
                        string variableType = variableTypes[i];
                        if(variableType == "enum")
                        {
                            variableType="int";
                        }
                        scriptCodeBuilder.AppendLine($"            instance.{variableName} = ({variableType})Convert.ChangeType(rowElement.Element(\"{variableName}\")?.Value, typeof({variableType}));");
                    }

                    scriptCodeBuilder.AppendLine();
                    if (hasThirdKey)
                    {
                        // List에 추가하는 코드
                        scriptCodeBuilder.AppendLine("            if (!dictionary.ContainsKey(instance.index))");
                        scriptCodeBuilder.AppendLine("            {");
                        scriptCodeBuilder.AppendLine($"                dictionary[instance.index] = new List<{className}>();");
                        scriptCodeBuilder.AppendLine("            }");
                        scriptCodeBuilder.AppendLine($"            dictionary[instance.index].Add(instance);");
                    }
                    else
                    {
                        scriptCodeBuilder.AppendLine("            dictionary.Add(instance.index, instance);"); // 'index'를 키로 사용
                    }
                    scriptCodeBuilder.AppendLine("        }");
                    scriptCodeBuilder.AppendLine("        return dictionary;");
                    scriptCodeBuilder.AppendLine("    }");

                    scriptCodeBuilder.AppendLine("}"); // 클래스 닫기
                    scriptCodeBuilder.AppendLine(); // 시트 간 공백 추가
                }
            }

            // 완성된 스크립트 코드를 반환합니다.
            return scriptCodeBuilder.ToString();
        }

        public static void GeneratedEnum(string excelFilePath, Dictionary<string, Dictionary<string, int>> enumDictionary)
        {
            // 엑셀 파일을 엽니다.
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {

                    if (worksheet.Name.Contains("_sub", StringComparison.OrdinalIgnoreCase))
                    {
                        //[W] 서브 테이블인 경우는 스크립트를 생성하지 않음.
                        continue;
                    }

                    var usedCells = worksheet.Row(1).CellsUsed().ToList(); // 첫 번째 행의 사용된 셀들

                    // 시트명과 대응되는 Dictionary 생성
                    var sheetEnumDict = new Dictionary<string, int>();

                    // 열을 순환하면서 조건 확인
                    for (int i = 0; i < usedCells.Count; i++)
                    {
                        // 2번 행의 n열이 "ignore"이면 continue
                        var secondRowValue = worksheet.Row(2).Cell(i + 1).GetValue<string>();
                        if (secondRowValue.Equals("ignore", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // 3번 행의 n열에 값이 있으면 continue
                        var thirdRowValue = worksheet.Row(3).Cell(i + 1).GetValue<string>();
                        if (!string.IsNullOrEmpty(thirdRowValue))
                        {
                            continue;
                        }

                        // 4번째 행이 "enum"이 아니면 continue
                        var fourthRowValue = worksheet.Row(4).Cell(i + 1).GetValue<string>();
                        if (!fourthRowValue.Equals("enum", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        // 5번째 행부터 데이터를 읽어들입니다.
                        int valueCounter = 0;
                        for (int rowIdx = 5; rowIdx <= worksheet.LastRowUsed().RowNumber(); rowIdx++)
                        {
                            var row = worksheet.Row(rowIdx);
                            string enumString = row.Cell(i + 1).GetValue<string>(); // 현재 열의 값을 사용

                            if (!string.IsNullOrEmpty(enumString) && !sheetEnumDict.ContainsKey(enumString))
                            {
                                sheetEnumDict.Add(enumString, valueCounter);
                                valueCounter++;
                            }
                        }
                    }

                    // 최종적으로 enumDictionary에 추가
                    if (sheetEnumDict.Count > 0)
                    {
                        enumDictionary[worksheet.Name] = sheetEnumDict;
                    }
                }
            }
        }
    }
}
