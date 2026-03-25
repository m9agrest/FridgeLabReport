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
            Time,

            // Схемные датчики
            T1, T2, T3, T4, T5, T6, T7, T8, T9, T10,
            T11, T12, T13, T14, T15, T16, T17, T18, T19, T20,
            T21, T22, T23, T24, T25, T26, T27,

            // Камера
            ChamberTemperature, // Температура в камере
            ChamberHumidity, // Влажность в камере

            // Процессные каналы
            Pc, // Давление конденсации
            Pe, // Давление кипения
            TcFilter, // Температура на фильтре
            TeSuction, // Температура на всасывании
            TCompressor, // Температура компрессора
            TCondInAir, // Вход в конденсатор (воздух)
            TCondOutAir, // Выход из конденсатора (воздух)
            TEvapInAir, // Вход в испаритель (воздух)
            TEvapOutAir, // Выход из испарителя (воздух)

            // Электрика
            Voltage, // Напряжение
            Current, // Ток
            Frequency, // Частота
            Power, // Мощность

            // Дополнительно
            HeaterPower2, // Мощность второго нагревателя
            DefrostPower, // Мощность оттайки
            DefrostTemperature1, // Температура датчика оттайки 1
            DefrostTemperature2, // Температура датчика оттайки 2
        }
        //считаем с 0 линии
        private const int lineTime = 1;//поле с временем. пример: "2026_02_05_10_19_17", есть еще 3853131557.882607, равноценно ли?
        private const int lineHeader = 5;//поле с шапкой колонок. пример: "time","C0","C1",...
        private const int lineData = 7;//с какой строки начинаются сами данные. пример: 430365.77607, n, n, ..., необходимо, так как шапка дублируется


        private static readonly HashSet<string> trashData = new() { "n", "N", "NaN", "nan", "NaT", "nat", "", "9999", "-9999", "0" };//мусорные данные, которые необходимо отсекать при парсинге данных в double


        private readonly long time = 0;
        private readonly Dictionary<string, int> ChannelMap = new();
        private static readonly Dictionary<DataField, string> FieldToChannel = new()
        {
            { DataField.Time, "time" },

            // Схемные датчики
            { DataField.T1, "C310" },
            { DataField.T2, "C311" },
            { DataField.T3, "C320" },
            { DataField.T4, "C321" },
            { DataField.T5, "C330" },
            { DataField.T6, "C331" },
            { DataField.T7, "C340" },
            { DataField.T8, "C341" },
            { DataField.T9, "C342" },
            { DataField.T10, "C350" },

            { DataField.T11, "C351" },
            { DataField.T12, "C352" },
            { DataField.T13, "C353" },
            { DataField.T14, "C360" },
            { DataField.T15, "C361" },
            { DataField.T16, "C370" },
            { DataField.T17, "C371" },
            { DataField.T18, "C380" },
            { DataField.T19, "C381" },
            { DataField.T20, "C390" },

            { DataField.T21, "C391" },
            { DataField.T22, "C400" },
            { DataField.T23, "C401" },
            { DataField.T24, "C410" },
            { DataField.T25, "C411" },
            { DataField.T26, "C420" },
            { DataField.T27, "C421" },

            // Камера
            { DataField.ChamberTemperature, "C531" },
            { DataField.ChamberHumidity, "C532" },

            // Процессные каналы
            // Pc / Pe — наиболее вероятные кандидаты по данным
            { DataField.Pc, "C533" },
            { DataField.Pe, "C534" },

            // Эти каналы я привязываю как РАБОЧУЮ предварительную карту
            // потому что они похожи на отдельные процессные температуры,
            // а не на T1..T27 со схемы
            { DataField.TcFilter, "C550" },
            { DataField.TeSuction, "C551" },
            { DataField.TCompressor, "C552" },
            { DataField.TCondInAir, "C553" },
            { DataField.TCondOutAir, "C554" },
            { DataField.TEvapInAir, "C555" },
            { DataField.TEvapOutAir, "C623" },

            // Электрика
            { DataField.Voltage, "C460" },
            { DataField.Current, "C461" },
            { DataField.Frequency, "C463" },

            // Мощность — как рабочий основной канал
            { DataField.Power, "C559" },

            // Дополнительно
            { DataField.HeaterPower2, "C569" },
            { DataField.DefrostPower, "C628" },
            { DataField.DefrostTemperature1, "C626" },
            { DataField.DefrostTemperature2, "C629" }
        };

        private readonly List<DataRow> dataRows = new();
        public readonly IReadOnlyList<DataRow> DataRows;
        public IReadOnlyCollection<string> Titles => ChannelMap.Keys;
        private DataContainer(long time)//чтобы запретить внешний конструктор
        {
            
            this.time = time;
            DataRows = new ReadOnlyCollection<DataRow>(dataRows);
        }

        public bool IsPresetField(DataField field) => FieldToChannel.ContainsKey(field);
        public string GetField(DataField field) => FieldToChannel[field];


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
            public long LocalTime => time - DateTimeOffset.Now.ToUnixTimeMilliseconds();

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
