using FridgeLabReport.Data;
using Microsoft.Win32;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace FridgeLabReport
{
    public partial class MainWindow : Window
    {
        private DataContainer? dc;

        private long minTime;
        private long maxTime;
        private long selectedFrom;
        private long selectedTo;

        public MainWindow()
        {
            InitializeComponent();
            SetDisabledState();
        }

        private void SetDisabledState()
        {
            DpFromDate.IsEnabled = false;
            TbFromTime.IsEnabled = false;
            DpToDate.IsEnabled = false;
            TbToTime.IsEnabled = false;
            BtnApplyRange.IsEnabled = false;
            BtnBuildReport.IsEnabled = false;

            DpFromDate.SelectedDate = null;
            DpToDate.SelectedDate = null;
            TbFromTime.Text = "";
            TbToTime.Text = "";
            TbRangeLimit.Text = "";

            BindingsPanel.Children.Clear();
            TbStatus.Text = "Файл не выбран";
        }

        private void SetEnabledState()
        {
            DpFromDate.IsEnabled = true;
            TbFromTime.IsEnabled = true;
            DpToDate.IsEnabled = true;
            TbToTime.IsEnabled = true;
            BtnApplyRange.IsEnabled = true;
            BtnBuildReport.IsEnabled = true;
        }

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = "DAT files (*.dat)|*.dat|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
                return;

            BusyWindow busyWindow = new BusyWindow("Читаем файл...")
            {
                Owner = this
            };

            try
            {
                busyWindow.Show();

                DataContainer loadedDc = await Task.Run(() =>
                {
                    return DataContainer.GenerateFromPath(dialog.FileName);
                });

                if (loadedDc.DataRows.Count == 0)
                    throw new ArgumentException("После парсинга не найдено ни одной строки данных");

                dc = loadedDc;

                TbFilePath.Text = dialog.FileName;

                minTime = dc.DataRows[0].Time;
                maxTime = dc.DataRows[dc.DataRows.Count - 1].Time;

                selectedFrom = minTime;
                selectedTo = maxTime;

                FillDateTimeFields();
                RebuildBindings();
                SetEnabledState();

                TbStatus.Text = $"Загружено строк: {dc.DataRows.Count}";
            }
            catch (Exception ex)
            {
                dc = null;
                TbFilePath.Text = dialog.FileName;
                SetDisabledState();

                TbStatus.Text = "Ошибка загрузки";
                MessageBox.Show(ex.Message, "Ошибка парсинга", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                busyWindow.Close();
            }
        }

        private void FillDateTimeFields()
        {
            SetDateTimeToControls(DpFromDate, TbFromTime, minTime);
            SetDateTimeToControls(DpToDate, TbToTime, maxTime);

            DateTime minDt = FromUnixMs(minTime);
            DateTime maxDt = FromUnixMs(maxTime);

            DpFromDate.DisplayDateStart = minDt.Date;
            DpFromDate.DisplayDateEnd = maxDt.Date;
            DpToDate.DisplayDateStart = minDt.Date;
            DpToDate.DisplayDateEnd = maxDt.Date;

            TbRangeLimit.Text =
                $"Доступный диапазон: {minDt:dd.MM.yyyy HH:mm:ss.fff} — {maxDt:dd.MM.yyyy HH:mm:ss.fff}";
        }

        private void SetDateTimeToControls(DatePicker datePicker, TextBox timeBox, long unixMs)
        {
            DateTime dt = FromUnixMs(unixMs);
            datePicker.SelectedDate = dt.Date;
            timeBox.Text = dt.ToString("HH:mm:ss.fff");
        }

        private DateTime FromUnixMs(long unixMs)
        {
            return DateTimeOffset.FromUnixTimeMilliseconds(unixMs).LocalDateTime;
        }

        private bool TryReadDateTime(DatePicker datePicker, TextBox timeBox, out long unixMs)
        {
            unixMs = 0;

            if (datePicker.SelectedDate == null)
                return false;

            string rawTime = timeBox.Text.Trim();

            string[] formats =
            {
                @"hh\:mm\:ss",
                @"hh\:mm\:ss\.f",
                @"hh\:mm\:ss\.ff",
                @"hh\:mm\:ss\.fff"
            };

            if (!TimeSpan.TryParseExact(
                    rawTime,
                    formats,
                    CultureInfo.InvariantCulture,
                    out TimeSpan timePart))
            {
                return false;
            }

            DateTime dt = datePicker.SelectedDate.Value.Date + timePart;
            unixMs = new DateTimeOffset(dt).ToUnixTimeMilliseconds();
            return true;
        }

        private void RebuildBindings()
        {
            BindingsPanel.Children.Clear();

            if (dc == null)
                return;

            int tCount = GetSelectedTCount();

            AddBindingRow(DataContainer.DataField.Time, "Время");

            for (int i = 0; i < tCount; i++)
            {
                var field = (DataContainer.DataField)i + 1; // потому что Time теперь первый
                AddBindingRow(field, $"T{i + 1}");
            }

            AddBindingRow(DataContainer.DataField.ChamberTemperature, "Температура камеры");
            AddBindingRow(DataContainer.DataField.ChamberHumidity, "Влажность камеры");

            AddBindingRow(DataContainer.DataField.Pc, "Pc");
            AddBindingRow(DataContainer.DataField.Pe, "Pe");
            AddBindingRow(DataContainer.DataField.TcFilter, "Температура фильтра");
            AddBindingRow(DataContainer.DataField.TeSuction, "Температура всасывания");
            AddBindingRow(DataContainer.DataField.TCompressor, "Температура компрессора");
            AddBindingRow(DataContainer.DataField.TCondInAir, "Воздух на входе конденсатора");
            AddBindingRow(DataContainer.DataField.TCondOutAir, "Воздух на выходе конденсатора");
            AddBindingRow(DataContainer.DataField.TEvapInAir, "Воздух на входе испарителя");
            AddBindingRow(DataContainer.DataField.TEvapOutAir, "Воздух на выходе испарителя");

            AddBindingRow(DataContainer.DataField.Voltage, "Напряжение");
            AddBindingRow(DataContainer.DataField.Current, "Ток");
            AddBindingRow(DataContainer.DataField.Frequency, "Частота");
            AddBindingRow(DataContainer.DataField.Power, "Мощность");

            AddBindingRow(DataContainer.DataField.HeaterPower2, "Мощность нагревателя 2");
            AddBindingRow(DataContainer.DataField.DefrostPower, "Мощность оттайки");
            AddBindingRow(DataContainer.DataField.DefrostTemperature1, "Температура оттайки 1");
            AddBindingRow(DataContainer.DataField.DefrostTemperature2, "Температура оттайки 2");
        }

        private void AddBindingRow(DataContainer.DataField field, string name)
        {
            if (dc == null)
                return;

            Grid row = new Grid()
            {
                Margin = new Thickness(0, 0, 0, 6)
            };

            row.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(260) });
            row.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(220) });

            TextBlock title = new TextBlock()
            {
                Text = name,
                VerticalAlignment = VerticalAlignment.Center
            };

            ComboBox combo = new ComboBox()
            {
                ItemsSource = dc.Titles,
                Tag = field
            };

            if (dc.IsPresetField(field))
                combo.SelectedItem = dc.GetField(field);

            Grid.SetColumn(title, 0);
            Grid.SetColumn(combo, 1);

            row.Children.Add(title);
            row.Children.Add(combo);

            BindingsPanel.Children.Add(row);
        }

        private int GetSelectedTCount()
        {
            if (CbTCount.SelectedItem is ComboBoxItem item &&
                int.TryParse(item.Content?.ToString(), out int value))
            {
                return value;
            }

            if (int.TryParse(CbTCount.Text, out int value2))
                return value2;

            return 15;
        }

        private void CbTCount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dc == null)
                return;

            RebuildBindings();
        }

        private void BtnApplyRange_Click(object sender, RoutedEventArgs e)
        {
            if (dc == null)
                return;

            if (!TryReadDateTime(DpFromDate, TbFromTime, out long from))
            {
                MessageBox.Show("Некорректное время \"От\".\nФормат: HH:mm:ss.fff",
                    "Диапазон", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!TryReadDateTime(DpToDate, TbToTime, out long to))
            {
                MessageBox.Show("Некорректное время \"До\".\nФормат: HH:mm:ss.fff",
                    "Диапазон", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (from < minTime) from = minTime;
            if (from > maxTime) from = maxTime;
            if (to < minTime) to = minTime;
            if (to > maxTime) to = maxTime;

            if (from > to)
            {
                long temp = from;
                from = to;
                to = temp;
            }

            selectedFrom = from;
            selectedTo = to;

            SetDateTimeToControls(DpFromDate, TbFromTime, selectedFrom);
            SetDateTimeToControls(DpToDate, TbToTime, selectedTo);

            int count = 0;
            foreach (var row in dc.DataRows)
            {
                if (row.Time >= selectedFrom && row.Time <= selectedTo)
                    count++;
            }

            TbStatus.Text = $"Диапазон: {count} строк";
        }

        private async void BtnBuildReport_Click(object sender, RoutedEventArgs e)
        {
            if (dc == null)
                return;

            Dictionary<DataContainer.DataField, string> fields = new();

            foreach (var child in BindingsPanel.Children)
            {
                if (child is not Grid row || row.Children.Count < 2)
                    continue;

                if (row.Children[1] is not ComboBox combo)
                    continue;

                if (combo.Tag is not DataContainer.DataField field)
                    continue;

                if (combo.SelectedItem is not string channelName)
                    continue;

                fields[field] = channelName;
            }

            List<DataContainer.DataRow> dataRows = dc.DataRows
                .Where(x => x.Time >= selectedFrom && x.Time <= selectedTo)
                .ToList();

            int tCount = GetSelectedTCount();

            string sourceFileName = Path.GetFileNameWithoutExtension(TbFilePath.Text);
            if (string.IsNullOrWhiteSpace(sourceFileName))
                sourceFileName = "report";

            SaveFileDialog dialog = new SaveFileDialog()
            {
                Filter = "Excel file (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                FileName = sourceFileName + ".xlsx",
                OverwritePrompt = true
            };

            if (dialog.ShowDialog() != true)
                return;

            BusyWindow busyWindow = new BusyWindow("Генерируем файл...")
            {
                Owner = this
            };

            try
            {
                busyWindow.Show();

                await Task.Run(() =>
                {
                    Generator.GenerateXlsx(dialog.FileName, tCount, fields, dataRows);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Ошибка при генерации файла:\n" + ex.Message,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }
            finally
            {
                busyWindow.Close();
            }

            MessageBoxResult result = MessageBox.Show(this,
                "Файл успешно сгенерирован.\n\nОткрыть его?",
                "Готово",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (result == MessageBoxResult.Yes && File.Exists(dialog.FileName))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = dialog.FileName,
                    UseShellExecute = true
                });
            }
        }
    }
}