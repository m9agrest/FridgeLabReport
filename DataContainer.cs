using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;

namespace FridgeLabReport.Data
{
    class DataContainer
    {
        public enum DataField
        {
            T1, T2, T3, T4, T5, T6, T7, T8, T9, T10,
            T11, T12, T13, T14, T15, T16, T17, T18, T19, T20, T21,
            Time, Power
        }
        //считаем с 0 линии
        private const int lineTime = 1;//поле с временем. пример: "2026_02_05_10_19_17", есть еще 3853131557.882607, равноценно ли?
        private const int lineHeader = 5;//поле с шапкой колонок. пример: "time","C0","C1",...
        private const int lineData = 7;//с какой строки начинаются сами данные. пример: 430365.77607, n, n, ..., необходимо, так как шапка дублируется


        private static readonly HashSet<string> trashData = new() { "n", "N", "NaN", "nan", "NaT", "nat", "", "9999", "-9999", "1" };//мусорные данные, которые необходимо отсекать при парсинге данных в double


        private readonly long time = 0;
        private readonly Dictionary<string, int> ChannelMap = new();
        private static readonly Dictionary<DataField, string> FieldToChannel = new()
        {
            { DataField.Time, "time" }
        };

        private readonly List<DataRow> dataRows = new();
        public readonly IReadOnlyList<DataRow> DataRows;
        public IReadOnlyCollection<string> Titles => ChannelMap.Keys;
        private DataContainer(long time)//чтобы запретить внешний конструктор
        {
            
            this.time = time;
            DataRows = new ReadOnlyCollection<DataRow>(dataRows);
        }


        public static DataContainer GenerateFromPath(string path) => GenerateFromData(File.ReadAllText(path));
        public static DataContainer GenerateFromData(string data)
        {
            DataContainer dc;

            if (string.IsNullOrWhiteSpace(data))
                throw new ArgumentException("Файл с данными пустой");

            string[] lines = data
                            .Replace("\r\n", "\n")
                            .Replace('\r', '\n')
                            .Split('\n', StringSplitOptions.None);

            if (lines.Length < lineData)
                throw new ArgumentNullException("В файле нет данных");

            //читаем время на строке lineTime и переносим в dc.time
            string rawTime = lines[lineTime].Trim().Trim('"');
            if (DateTime.TryParseExact(rawTime, "yyyy_MM_dd_HH_mm_ss",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt))
            {
                dc = new DataContainer(new DateTimeOffset(dt).ToUnixTimeMilliseconds());
            }
            else
            {
                throw new ArgumentException("Не удалось распарсить время файла: " + rawTime);
            }

            //читаем шапка на строке lineHeader и заполняем dc.ChannelMap
            string[] header = lines[lineHeader].Split(',');
            for (int i = 0; i < header.Length; i++)
            {
                string name = header[i].Trim().Trim('"');

                if (string.IsNullOrWhiteSpace(name))
                    continue;

                dc.ChannelMap[name] = i;
            }

            if (!dc.ChannelMap.ContainsKey(FieldToChannel[DataField.Time]))
                throw new ArgumentException("В файле не найден канал time");

            //читаем данные начиная с lineData и заполняем AddDataRow, перед этим парся данные в double и отсеивая мусор из trashData
            for (int lineIndex = lineData; lineIndex < lines.Length; lineIndex++)
            {
                string line = lines[lineIndex];

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',');
                Dictionary<int, double> values = new();

                for (int i = 0; i < parts.Length; i++)
                {
                    string raw = parts[i].Trim().Trim('"');

                    if (trashData.Contains(raw))
                        continue;

                    if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                    {
                        values[i] = value;
                    }
                }

                if (!values.ContainsKey(dc.ChannelMap[FieldToChannel[DataField.Time]]))
                    continue;

                dc.dataRows.Add(new DataRow(dc, values));
            }


            return dc; 
        }



        public sealed class DataRow
        {
            private readonly DataContainer dc;
            private readonly long time;
            private readonly Dictionary<int, double> Values;

            public long Time => time + dc.time;

            internal DataRow(DataContainer dataContainer, Dictionary<int, double> values)
            {
                dc = dataContainer;
                Values = values;
                time = (long)this[DataField.Time];
            }

            public double this[int index] => GetValue(index);
            public double this[string channelName] => GetValue(channelName);
            public double this[DataField field]  => GetValue(field);

            public double GetValue(int index) => Values.GetValueOrDefault(index, default);
            public double GetValue(string channelName) => GetValue(dc.ChannelMap[channelName]);
            public double GetValue(DataField field) => GetValue(FieldToChannel[field]);
        }
    }
}
