public class Test1
{
    public int index { get; set; } = 0;
    public string test2 { get; set; } = "";
    public float test3 { get; set; } = 0.0f;
    public byte test4 { get; set; } = 0;
    public bool test5 { get; set; } = false;
    public bool test6 { get; set; } = false;

    public static Dictionary<int, Test1> CreateData(string xmlPath)
    {
        var dictionary = new Dictionary<int, Test1>();
        var document = System.Xml.Linq.XDocument.Load(xmlPath);
        foreach (var rowElement in document.Descendants("Row"))
        {
            var instance = new Test1();
            instance.index = (int)Convert.ChangeType(rowElement.Element("index")?.Value, typeof(int));
            instance.test2 = (string)Convert.ChangeType(rowElement.Element("test2")?.Value, typeof(string));
            instance.test3 = (float)Convert.ChangeType(rowElement.Element("test3")?.Value, typeof(float));
            instance.test4 = (byte)Convert.ChangeType(rowElement.Element("test4")?.Value, typeof(byte));
            instance.test5 = (bool)Convert.ChangeType(rowElement.Element("test5")?.Value, typeof(bool));
            instance.test6 = (bool)Convert.ChangeType(rowElement.Element("test6")?.Value, typeof(bool));

            dictionary.Add(instance.index, instance);
        }
        return dictionary;
    }
}

public class Test2
{
    public int index { get; set; } = 0;
    public int Fast { get; set; } = 0;
    public float Fast1 { get; set; } = 0.0f;
    public byte Fast2 { get; set; } = 0;
    public bool Fast3 { get; set; } = false;
    public bool Fast4 { get; set; } = false;

    public static Dictionary<int, Test2> CreateData(string xmlPath)
    {
        var dictionary = new Dictionary<int, Test2>();
        var document = System.Xml.Linq.XDocument.Load(xmlPath);
        foreach (var rowElement in document.Descendants("Row"))
        {
            var instance = new Test2();
            instance.index = (int)Convert.ChangeType(rowElement.Element("index")?.Value, typeof(int));
            instance.Fast = (int)Convert.ChangeType(rowElement.Element("Fast")?.Value, typeof(int));
            instance.Fast1 = (float)Convert.ChangeType(rowElement.Element("Fast1")?.Value, typeof(float));
            instance.Fast2 = (byte)Convert.ChangeType(rowElement.Element("Fast2")?.Value, typeof(byte));
            instance.Fast3 = (bool)Convert.ChangeType(rowElement.Element("Fast3")?.Value, typeof(bool));
            instance.Fast4 = (bool)Convert.ChangeType(rowElement.Element("Fast4")?.Value, typeof(bool));

            dictionary.Add(instance.index, instance);
        }
        return dictionary;
    }
}

public class ThirdKeyTest
{
    public int index { get; set; } = 0;
    public string Fast { get; set; } = "";
    public float Fast1 { get; set; } = 0.0f;
    public byte Fast2 { get; set; } = 0;
    public bool Fast3 { get; set; } = false;
    public bool Fast4 { get; set; } = false;

    public static Dictionary<int, List<ThirdKeyTest>> CreateData(string xmlPath)
    {
        var dictionary = new Dictionary<int, List<ThirdKeyTest>>();
        var document = System.Xml.Linq.XDocument.Load(xmlPath);
        foreach (var rowElement in document.Descendants("Row"))
        {
            var instance = new ThirdKeyTest();
            instance.index = (int)Convert.ChangeType(rowElement.Element("index")?.Value, typeof(int));
            instance.Fast = (string)Convert.ChangeType(rowElement.Element("Fast")?.Value, typeof(string));
            instance.Fast1 = (float)Convert.ChangeType(rowElement.Element("Fast1")?.Value, typeof(float));
            instance.Fast2 = (byte)Convert.ChangeType(rowElement.Element("Fast2")?.Value, typeof(byte));
            instance.Fast3 = (bool)Convert.ChangeType(rowElement.Element("Fast3")?.Value, typeof(bool));
            instance.Fast4 = (bool)Convert.ChangeType(rowElement.Element("Fast4")?.Value, typeof(bool));

            if (!dictionary.ContainsKey(instance.index))
            {
                dictionary[instance.index] = new List<ThirdKeyTest>();
            }
            dictionary[instance.index].Add(instance);
        }
        return dictionary;
    }
}

public class Sheet1
{
    public int index { get; set; } = 0;
    public string Fast { get; set; } = "";
    public float Fast1 { get; set; } = 0.0f;
    public byte Fast2 { get; set; } = 0;
    public bool Fast3 { get; set; } = false;
    public bool Fast4 { get; set; } = false;

    public static Dictionary<int, Sheet1> CreateData(string xmlPath)
    {
        var dictionary = new Dictionary<int, Sheet1>();
        var document = System.Xml.Linq.XDocument.Load(xmlPath);
        foreach (var rowElement in document.Descendants("Row"))
        {
            var instance = new Sheet1();
            instance.index = (int)Convert.ChangeType(rowElement.Element("index")?.Value, typeof(int));
            instance.Fast = (string)Convert.ChangeType(rowElement.Element("Fast")?.Value, typeof(string));
            instance.Fast1 = (float)Convert.ChangeType(rowElement.Element("Fast1")?.Value, typeof(float));
            instance.Fast2 = (byte)Convert.ChangeType(rowElement.Element("Fast2")?.Value, typeof(byte));
            instance.Fast3 = (bool)Convert.ChangeType(rowElement.Element("Fast3")?.Value, typeof(bool));
            instance.Fast4 = (bool)Convert.ChangeType(rowElement.Element("Fast4")?.Value, typeof(bool));

            dictionary.Add(instance.index, instance);
        }
        return dictionary;
    }
}


