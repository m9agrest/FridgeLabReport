using ClosedXML.Excel;
using FridgeLabReport.Data;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace FridgeLabReport
{
    public partial class MainWindow : Window
    {
        private const int DefaultTCount = 15;

        private DataContainer? dc;

        private long minTime;
        private long maxTime;
        private long selectedFrom;
        private long selectedTo;

        private ReportSettings reportSettings = new();

        private static readonly JsonSerializerOptions ReportSettingsJsonOptions = new()
        {
            WriteIndented = true
        };

        private readonly string reportSettingsFilePath = Path.Combine(
            AppContext.BaseDirectory,
            "conf",
            "report_settings.json");

        private readonly string confDirPath = Path.Combine(AppContext.BaseDirectory, "conf");
        private readonly string channelConfigPath = Path.Combine(AppContext.BaseDirectory, "conf", "channel_bindings.json");

        private readonly Dictionary<DataContainer.DataField, string> defaultChannelBindings = new();

        private bool isChannelConfigDirty;
        private bool isApplyingChannelConfig;

        private sealed class ChannelConfig
        {
            public int TCount { get; set; } = DefaultTCount;
            public Dictionary<string, string> Bindings { get; set; } = new();
        }

        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;
            LoadDefaultReportSettings();
            SetDisabledState();
        }

        private void LoadDefaultReportSettings()
        {
            try
            {
                if (!File.Exists(reportSettingsFilePath))
                {
                    reportSettings = new ReportSettings();
                    return;
                }

                string json = File.ReadAllText(reportSettingsFilePath);
                ReportSettings? loaded = JsonSerializer.Deserialize<ReportSettings>(json, ReportSettingsJsonOptions);

                reportSettings = loaded ?? new ReportSettings();
            }
            catch
            {
                reportSettings = new ReportSettings();
            }
        }

        private void SetDisabledState()
        {
            DpFromDate.IsEnabled = false;
            TbFromTime.IsEnabled = false;
            DpToDate.IsEnabled = false;
            TbToTime.IsEnabled = false;
            BtnApplyRange.IsEnabled = false;
            BtnBuildReport.IsEnabled = false;
            BtnSaveChannelConfig.IsEnabled = false;
            BtnResetChannelConfig.IsEnabled = false;

            DpFromDate.SelectedDate = null;
            DpToDate.SelectedDate = null;
            TbFromTime.Text = string.Empty;
            TbToTime.Text = string.Empty;
            TbRangeLimit.Text = string.Empty;

            BindingsPanel.Children.Clear();
            TbStatus.Text = "Файл не выбран";

            defaultChannelBindings.Clear();
            isChannelConfigDirty = false;
            UpdateReportSettingsSummary();
        }

        private void SetEnabledState()
        {
            DpFromDate.IsEnabled = true;
            TbFromTime.IsEnabled = true;
            DpToDate.IsEnabled = true;
            TbToTime.IsEnabled = true;
            BtnApplyRange.IsEnabled = true;
            BtnBuildReport.IsEnabled = true;
            BtnSaveChannelConfig.IsEnabled = true;
            BtnResetChannelConfig.IsEnabled = true;
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

                DataContainer loadedDc = await Task.Run(() => DataContainer.GenerateFromPath(dialog.FileName));

                if (loadedDc.DataRows.Count == 0)
                    throw new ArgumentException("После парсинга не найдено ни одной строки данных");

                ChannelConfig? savedConfig = LoadChannelConfig();
                SetSelectedTCount(savedConfig?.TCount ?? DefaultTCount);

                dc = loadedDc;
                CaptureDefaultChannelBindings();

                TbFilePath.Text = dialog.FileName;

                minTime = dc.DataRows[0].Time;
                maxTime = dc.DataRows[dc.DataRows.Count - 1].Time;

                selectedFrom = minTime;
                selectedTo = maxTime;

                FillDateTimeFields();
                RebuildBindings(keepCurrentBindings: false);
                ApplyChannelConfig(savedConfig);
                SetEnabledState();

                isChannelConfigDirty = false;
                TbStatus.Text = $"Загружено строк: {dc.DataRows.Count}";
            }
            catch (Exception ex)
            {
                dc = null;
                TbFilePath.Text = dialog.FileName;
                SetDisabledState();

                TbStatus.Text = "Ошибка загрузки";
                MessageBox.Show(this, ex.Message, "Ошибка парсинга", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                busyWindow.Close();
            }
        }

        private void UpdateReportSettingsSummary()
        {
            string lab = string.IsNullOrWhiteSpace(reportSettings.LabAssistantFullName)
                ? "—"
                : reportSettings.LabAssistantFullName;

            string test = string.IsNullOrWhiteSpace(reportSettings.TestName)
                ? "—"
                : reportSettings.TestName;

            string minPower = reportSettings.MinPowerHighlight?.ToString(CultureInfo.CurrentCulture) ?? "—";
            string minTCompressor = reportSettings.MinTCompressorHighlight?.ToString(CultureInfo.CurrentCulture) ?? "—";
            string maxAllT = reportSettings.MaxAllT?.ToString(CultureInfo.CurrentCulture) ?? "—";

            TbReportSettingsSummary.Text =
                $"Лаборант: {lab}; испытание: {test}; мин. мощность: {minPower}; мин. Tcompr: {minTCompressor}; макс. всех T: {maxAllT}";
        }

        private void BtnReportSettings_Click(object sender, RoutedEventArgs e)
        {
            ReportSettingsWindow window = new ReportSettingsWindow(reportSettings)
            {
                Owner = this
            };

            if (window.ShowDialog() == true)
            {
                reportSettings = window.ResultSettings;
                UpdateReportSettingsSummary();
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

        private void CaptureDefaultChannelBindings()
        {
            defaultChannelBindings.Clear();

            if (dc == null)
                return;

            foreach (DataContainer.DataField field in Enum.GetValues<DataContainer.DataField>())
            {
                if (!dc.IsPresetField(field))
                    continue;

                string channel = dc.GetField(field);
                if (!string.IsNullOrWhiteSpace(channel))
                    defaultChannelBindings[field] = channel;
            }
        }

        private void RebuildBindings(bool keepCurrentBindings = true)
        {
            Dictionary<DataContainer.DataField, string> currentBindings =
                keepCurrentBindings
                    ? GetBindingsFromUi()
                    : new Dictionary<DataContainer.DataField, string>();

            BindingsPanel.Children.Clear();

            if (dc == null)
                return;

            int tCount = GetSelectedTCount();

            for (int i = 0; i < tCount; i++)
            {
                DataContainer.DataField field = (DataContainer.DataField)i + 1;
                AddBindingRow(field, $"T{i + 1}", currentBindings);
            }

            AddBindingRow(DataContainer.DataField.ChamberTemperature, "Температура камеры", currentBindings);
            AddBindingRow(DataContainer.DataField.ChamberHumidity, "Влажность камеры", currentBindings);

            AddBindingRow(DataContainer.DataField.Pc, "Pc", currentBindings);
            AddBindingRow(DataContainer.DataField.Pe, "Pe", currentBindings);
            AddBindingRow(DataContainer.DataField.TcFilter, "Температура фильтра", currentBindings);
            AddBindingRow(DataContainer.DataField.TeSuction, "Температура всасывания", currentBindings);
            AddBindingRow(DataContainer.DataField.TCompressor, "Температура компрессора", currentBindings);
            AddBindingRow(DataContainer.DataField.TCondInAir, "Воздух на входе конденсатора", currentBindings);
            AddBindingRow(DataContainer.DataField.TCondOutAir, "Воздух на выходе конденсатора", currentBindings);
            AddBindingRow(DataContainer.DataField.TEvapInAir, "Воздух на входе испарителя", currentBindings);
            AddBindingRow(DataContainer.DataField.TEvapOutAir, "Воздух на выходе испарителя", currentBindings);

            AddBindingRow(DataContainer.DataField.Voltage, "Напряжение", currentBindings);
            AddBindingRow(DataContainer.DataField.Current, "Ток", currentBindings);
            AddBindingRow(DataContainer.DataField.Frequency, "Частота", currentBindings);
            AddBindingRow(DataContainer.DataField.Power, "Мощность", currentBindings);

            AddBindingRow(DataContainer.DataField.HeaterPower2, "Мощность нагревателя 2", currentBindings);
            AddBindingRow(DataContainer.DataField.DefrostPower, "Мощность оттайки", currentBindings);
            AddBindingRow(DataContainer.DataField.DefrostTemperature1, "Температура оттайки 1", currentBindings);
            AddBindingRow(DataContainer.DataField.DefrostTemperature2, "Температура оттайки 2", currentBindings);
        }

        private void AddBindingRow(
            DataContainer.DataField field,
            string name,
            Dictionary<DataContainer.DataField, string> currentBindings)
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
            combo.SelectionChanged += BindingCombo_SelectionChanged;

            if (currentBindings.TryGetValue(field, out string? currentChannel) &&
                !string.IsNullOrWhiteSpace(currentChannel) &&
                dc.Titles.Contains(currentChannel))
            {
                combo.SelectedItem = currentChannel;
            }
            else if (dc.IsPresetField(field))
            {
                string presetChannel = dc.GetField(field);
                if (dc.Titles.Contains(presetChannel))
                    combo.SelectedItem = presetChannel;
            }

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

            return DefaultTCount;
        }

        private void SetSelectedTCount(int tCount)
        {
            isApplyingChannelConfig = true;
            try
            {
                foreach (object item in CbTCount.Items)
                {
                    if (item is ComboBoxItem comboBoxItem &&
                        int.TryParse(comboBoxItem.Content?.ToString(), out int value) &&
                        value == tCount)
                    {
                        CbTCount.SelectedItem = comboBoxItem;
                        return;
                    }
                }

                CbTCount.SelectedIndex = 0;
            }
            finally
            {
                isApplyingChannelConfig = false;
            }
        }

        private Dictionary<DataContainer.DataField, string> GetBindingsFromUi()
        {
            Dictionary<DataContainer.DataField, string> fields = new();

            foreach (object child in BindingsPanel.Children)
            {
                if (child is not Grid row || row.Children.Count < 2)
                    continue;

                if (row.Children[1] is not ComboBox combo)
                    continue;

                if (combo.Tag is not DataContainer.DataField field)
                    continue;

                if (combo.SelectedItem is not string channelName || string.IsNullOrWhiteSpace(channelName))
                    continue;

                fields[field] = channelName;
            }

            return fields;
        }

        private ChannelConfig? LoadChannelConfig()
        {
            try
            {
                if (!File.Exists(channelConfigPath))
                    return null;

                string json = File.ReadAllText(channelConfigPath);
                if (string.IsNullOrWhiteSpace(json))
                    return null;

                return JsonSerializer.Deserialize<ChannelConfig>(json);
            }
            catch
            {
                return null;
            }
        }

        private void ApplyChannelConfig(ChannelConfig? config)
        {
            if (dc == null || config?.Bindings == null)
                return;

            isApplyingChannelConfig = true;
            try
            {
                foreach (object child in BindingsPanel.Children)
                {
                    if (child is not Grid row || row.Children.Count < 2)
                        continue;

                    if (row.Children[1] is not ComboBox combo)
                        continue;

                    if (combo.Tag is not DataContainer.DataField field)
                        continue;

                    if (!config.Bindings.TryGetValue(field.ToString(), out string? channelName))
                        continue;

                    if (!string.IsNullOrWhiteSpace(channelName) && dc.Titles.Contains(channelName))
                        combo.SelectedItem = channelName;
                }
            }
            finally
            {
                isApplyingChannelConfig = false;
            }
        }

        private void ApplyBindings(Dictionary<DataContainer.DataField, string> bindings)
        {
            if (dc == null)
                return;

            isApplyingChannelConfig = true;
            try
            {
                foreach (object child in BindingsPanel.Children)
                {
                    if (child is not Grid row || row.Children.Count < 2)
                        continue;

                    if (row.Children[1] is not ComboBox combo)
                        continue;

                    if (combo.Tag is not DataContainer.DataField field)
                        continue;

                    if (bindings.TryGetValue(field, out string? channelName) &&
                        !string.IsNullOrWhiteSpace(channelName) &&
                        dc.Titles.Contains(channelName))
                    {
                        combo.SelectedItem = channelName;
                    }
                    else
                    {
                        combo.SelectedItem = null;
                    }
                }
            }
            finally
            {
                isApplyingChannelConfig = false;
            }
        }

        private bool SaveChannelConfig(bool showSuccessMessage)
        {
            if (dc == null)
            {
                MessageBox.Show(this,
                    "Сначала загрузите DAT-файл.",
                    "Сохранение настроек",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return false;
            }

            try
            {
                Directory.CreateDirectory(confDirPath);

                ChannelConfig config = new ChannelConfig
                {
                    TCount = GetSelectedTCount(),
                    Bindings = GetBindingsFromUi().ToDictionary(x => x.Key.ToString(), x => x.Value)
                };

                JsonSerializerOptions options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };

                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(channelConfigPath, json);

                isChannelConfigDirty = false;
                TbStatus.Text = "Настройки каналов сохранены";

                if (showSuccessMessage)
                {
                    MessageBox.Show(this,
                        "Настройки сохранены",
                        "Сохранение настроек",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Не удалось сохранить настройки:\n" + ex.Message,
                    "Сохранение настроек",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return false;
            }
        }

        private void MarkChannelConfigDirty()
        {
            if (isApplyingChannelConfig)
                return;

            isChannelConfigDirty = true;
        }

        private void BindingCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MarkChannelConfigDirty();
        }

        private void CbTCount_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dc == null)
                return;

            RebuildBindings();

            if (!isApplyingChannelConfig)
                MarkChannelConfigDirty();
        }

        private void BtnSaveChannelConfig_Click(object sender, RoutedEventArgs e)
        {
            SaveChannelConfig(showSuccessMessage: true);
        }

        private void BtnResetChannelConfig_Click(object sender, RoutedEventArgs e)
        {
            if (dc == null)
            {
                MessageBox.Show(this,
                    "Сначала загрузите DAT-файл.",
                    "Сброс настроек",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            MessageBoxResult result = MessageBox.Show(this,
                "Сбросить настройки каналов к значениям по умолчанию и удалить сохранённый конфиг?",
                "Сброс настроек",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                if (File.Exists(channelConfigPath))
                    File.Delete(channelConfigPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Не удалось удалить сохранённый конфиг:\n" + ex.Message,
                    "Сброс настроек",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            SetSelectedTCount(DefaultTCount);
            RebuildBindings(keepCurrentBindings: false);
            ApplyBindings(defaultChannelBindings);
            dc.SetFieldToChannel(new Dictionary<DataContainer.DataField, string>(defaultChannelBindings));

            isChannelConfigDirty = false;
            TbStatus.Text = "Настройки каналов сброшены";
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            if (!isChannelConfigDirty)
                return;

            MessageBoxResult result = MessageBox.Show(this,
                "Сохранить настройки выбора каналов перед закрытием?",
                "Сохранение настроек",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == MessageBoxResult.Yes && !SaveChannelConfig(showSuccessMessage: false))
                e.Cancel = true;
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
            foreach (DataContainer.DataRow row in dc.DataRows)
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

            Dictionary<DataContainer.DataField, string> fields = GetBindingsFromUi();

            List<DataContainer.DataRow> dataRows = dc.DataRows
                .Where(x => x.Time >= selectedFrom && x.Time <= selectedTo)
                .ToList();

            int tCount = GetSelectedTCount();

            string sourceFileName = Path.GetFileNameWithoutExtension(TbFilePath.Text);
            if (string.IsNullOrWhiteSpace(sourceFileName))
                sourceFileName = "report";

            SaveFileDialog xlsxDialog = new SaveFileDialog()
            {
                Filter = "Excel file (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                AddExtension = true,
                FileName = sourceFileName + ".xlsx",
                OverwritePrompt = true
            };

            if (xlsxDialog.ShowDialog() != true)
                return;

            string xlsxPath = xlsxDialog.FileName;

            SaveFileDialog docxDialog = new SaveFileDialog()
            {
                Filter = "Word file (*.docx)|*.docx",
                DefaultExt = "docx",
                AddExtension = true,
                FileName = sourceFileName + ".docx",
                OverwritePrompt = true
            };

            string? docxPath = null;
            if (docxDialog.ShowDialog() == true)
                docxPath = docxDialog.FileName;

            BusyWindow busyWindow = new BusyWindow("Генерируем файл(ы)...")
            {
                Owner = this
            };

            try
            {
                busyWindow.Show();

                await Task.Run(() =>
                {
                    dc.SetFieldToChannel(fields);

                    Generator.GenerateXlsx(xlsxPath, tCount, dataRows, reportSettings);

                    if (string.IsNullOrWhiteSpace(docxPath))
                        return;

                    using var wb = new XLWorkbook(xlsxPath);
                    IXLWorksheet ws = wb.Worksheet(1);

                    Generator.GenerateDocx(docxPath, tCount, dataRows, ws, reportSettings);
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

            string successText = string.IsNullOrWhiteSpace(docxPath)
                ? "Таблица успешно сгенерирована.\n\nОткрыть её?"
                : "Файлы успешно сгенерированы.\n\nОткрыть их?";

            MessageBoxResult result = MessageBox.Show(this,
                successText,
                "Готово",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (result == MessageBoxResult.Yes && File.Exists(xlsxPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = xlsxPath,
                    UseShellExecute = true
                });

                if (!string.IsNullOrWhiteSpace(docxPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = docxPath,
                        UseShellExecute = true
                    });
                }
            }
        }
    }
}
