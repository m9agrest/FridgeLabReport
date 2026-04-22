using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace FridgeLabReport
{
    public partial class ReportSettingsWindow : Window
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        private readonly string settingsFilePath;

        public ReportSettings ResultSettings { get; private set; }

        public ReportSettingsWindow(ReportSettings source)
        {
            InitializeComponent();

            settingsFilePath = Path.Combine(
                AppContext.BaseDirectory,
                "conf",
                "report_settings.json");

            ResultSettings = source.Clone();
            FillControls(ResultSettings);
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!TryBuildSettingsFromControls(out ReportSettings? settings))
                return;

            ResultSettings = settings;
            DialogResult = true;
            Close();
        }

        private void MiSaveDefaults_Click(object sender, RoutedEventArgs e)
        {
            if (!TryBuildSettingsFromControls(out ReportSettings? settings))
                return;

            try
            {
                string? directory = Path.GetDirectoryName(settingsFilePath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);

                string json = JsonSerializer.Serialize(settings, JsonOptions);
                File.WriteAllText(settingsFilePath, json);

                MessageBox.Show(this,
                    "Настройки по умолчанию сохранены.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Не удалось сохранить настройки по умолчанию.\n\n" + ex.Message,
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void MiResetDefaults_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(settingsFilePath))
                    File.Delete(settingsFilePath);

                FillControls(new ReportSettings());

                MessageBox.Show(this,
                    "Сохранённые настройки удалены. Поля сброшены.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "Не удалось сбросить сохранённые настройки.\n\n" + ex.Message,
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private bool TryBuildSettingsFromControls(out ReportSettings? settings)
        {
            settings = null;

            if (!TryParseNullableDouble(TbMinPower.Text, out double? minPower))
            {
                MessageBox.Show(this,
                    "Минимальная мощность введена некорректно.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }

            if (!TryParseNullableDouble(TbMinTCompressor.Text, out double? minTCompressor))
            {
                MessageBox.Show(this,
                    "Минимальная Tcompr введена некорректно.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }

            if (!TryParseNullableDouble(TbMinAllT.Text, out double? minAllT))
            {
                MessageBox.Show(this,
                    "Переход всех T за указанное значение введён некорректно.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }

            settings = new ReportSettings
            {
                LabAssistantFullName = TbLabAssistant.Text.Trim(),
                TestName = TbTestName.Text.Trim(),
                MinPowerHighlight = minPower,
                MinTCompressorHighlight = minTCompressor,
                MinAllT = minAllT,
            };

            return true;
        }

        private void FillControls(ReportSettings settings)
        {
            TbLabAssistant.Text = settings.LabAssistantFullName;
            TbTestName.Text = settings.TestName;
            TbMinPower.Text = FormatNullableDouble(settings.MinPowerHighlight);
            TbMinTCompressor.Text = FormatNullableDouble(settings.MinTCompressorHighlight);
            TbMinAllT.Text = FormatNullableDouble(settings.MinAllT);
        }

        private static bool TryParseNullableDouble(string raw, out double? value)
        {
            value = null;
            string text = raw.Trim();

            if (string.IsNullOrWhiteSpace(text))
                return true;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out double parsed) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        private static string FormatNullableDouble(double? value)
        {
            return value?.ToString(CultureInfo.CurrentCulture) ?? string.Empty;
        }
    }
}
