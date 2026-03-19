using FridgeLabReport.Data;
using Microsoft.Win32;
using System.Globalization;
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

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                Filter = "DAT files (*.dat)|*.dat|All files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                dc = DataContainer.GenerateFromPath(dialog.FileName);

                if (dc.DataRows.Count == 0)
                    throw new ArgumentException("После парсинга не найдено ни одной строки данных");

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

            for (int i = 0; i < tCount; i++)
                AddBindingRow((DataContainer.DataField)i);

            AddBindingRow(DataContainer.DataField.Power);
        }

        private void AddBindingRow(DataContainer.DataField field)
        {
            if (dc == null)
                return;

            Grid row = new Grid()
            {
                Margin = new Thickness(0, 0, 0, 6)
            };

            row.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(140) });
            row.ColumnDefinitions.Add(new ColumnDefinition() { Width = new GridLength(220) });

            TextBlock title = new TextBlock()
            {
                Text = field.ToString(),
                VerticalAlignment = VerticalAlignment.Center
            };

            ComboBox combo = new ComboBox()
            {
                ItemsSource = dc.Titles
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

        private void BtnBuildReport_Click(object sender, RoutedEventArgs e)
        {
            if (dc == null)
                return;

            string text =
                $"От: {FromUnixMs(selectedFrom):dd.MM.yyyy HH:mm:ss.fff}\n" +
                $"До: {FromUnixMs(selectedTo):dd.MM.yyyy HH:mm:ss.fff}\n\n" +
                $"Привязки:\n";

            foreach (var child in BindingsPanel.Children)
            {
                if (child is not Grid row || row.Children.Count < 2)
                    continue;

                if (row.Children[0] is not TextBlock title)
                    continue;

                if (row.Children[1] is not ComboBox combo)
                    continue;

                string channel = combo.SelectedItem?.ToString() ?? "(не выбрано)";
                text += $"{title.Text} -> {channel}\n";
            }

            MessageBox.Show(text, "Заглушка отчёта", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}