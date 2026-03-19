using FridgeLabReport.Data;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;

namespace FridgeLabReport
{
    public partial class MainWindow : Window
    {
        private DataContainer? dc;

        public MainWindow()
        {
            InitializeComponent();
            SetDisabledState();
        }

        private void SetDisabledState()
        {
            CbFrom.IsEnabled = false;
            CbTo.IsEnabled = false;
            BtnApplyRange.IsEnabled = false;
            BtnBuildReport.IsEnabled = false;

            CbFrom.ItemsSource = null;
            CbTo.ItemsSource = null;
            BindingsPanel.Children.Clear();

            TbStatus.Text = "Файл не выбран";
        }

        private void SetEnabledState()
        {
            CbFrom.IsEnabled = true;
            CbTo.IsEnabled = true;
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

                TbFilePath.Text = dialog.FileName;
                TbStatus.Text = $"Загружено строк: {dc.DataRows.Count}";

                FillRangeBoxes();
                RebuildBindings();

                SetEnabledState();
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

        private void FillRangeBoxes()
        {
            if (dc == null || dc.DataRows.Count == 0)
                return;

            CbFrom.ItemsSource = dc.DataRows;
            CbTo.ItemsSource = dc.DataRows;

            CbFrom.DisplayMemberPath = "Time";
            CbTo.DisplayMemberPath = "Time";

            CbFrom.SelectedIndex = 0;
            CbTo.SelectedIndex = dc.DataRows.Count - 1;
        }

        private void RebuildBindings()
        {
            BindingsPanel.Children.Clear();

            if (dc == null)
                return;

            int tCount = GetSelectedTCount();

            for (int i = 0; i < tCount; i++)
            {
                AddBindingRow((DataContainer.DataField)i);
            }

            AddBindingRow(DataContainer.DataField.Power);
        }

        private void AddBindingRow(DataContainer.DataField field)
        {
            if (dc == null)
                return;

            Grid row = new Grid()
            {
                Margin = new Thickness(0, 0, 0, 6),
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
                ItemsSource = dc.Titles,
                Tag = field
            };

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

            if (CbFrom.SelectedItem is not DataContainer.DataRow rowFrom ||
                CbTo.SelectedItem is not DataContainer.DataRow rowTo)
            {
                MessageBox.Show("Выбери диапазон.", "Диапазон", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            long from = Math.Min(rowFrom.Time, rowTo.Time);
            long to = Math.Max(rowFrom.Time, rowTo.Time);

            int count = 0;
            foreach (var row in dc.DataRows)
            {
                if (row.Time >= from && row.Time <= to)
                    count++;
            }

            TbStatus.Text = $"Диапазон: {count} строк";
        }

        private void BtnBuildReport_Click(object sender, RoutedEventArgs e)
        {
            if (dc == null)
                return;

            if (CbFrom.SelectedItem is not DataContainer.DataRow rowFrom ||
                CbTo.SelectedItem is not DataContainer.DataRow rowTo)
            {
                MessageBox.Show("Выбери диапазон.", "Диапазон", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            long from = Math.Min(rowFrom.Time, rowTo.Time);
            long to = Math.Max(rowFrom.Time, rowTo.Time);

            string text = $"От: {from}\nДо: {to}\n\nПривязки:\n";

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