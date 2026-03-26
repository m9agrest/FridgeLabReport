using System.Globalization;
using System.Windows;

namespace FridgeLabReport
{
    public partial class ReportSettingsWindow : Window
    {
        public ReportSettings ResultSettings { get; }

        public ReportSettingsWindow(ReportSettings source)
        {
            InitializeComponent();

            ResultSettings = source.Clone();

            TbLabAssistant.Text = ResultSettings.LabAssistantFullName;
            TbTestName.Text = ResultSettings.TestName;
            TbMinPower.Text = FormatNullableDouble(ResultSettings.MinPowerHighlight);
            TbMinTCompressor.Text = FormatNullableDouble(ResultSettings.MinTCompressorHighlight);
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!TryParseNullableDouble(TbMinPower.Text, out double? minPower))
            {
                MessageBox.Show(this,
                    "Минимальная мощность введена некорректно.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (!TryParseNullableDouble(TbMinTCompressor.Text, out double? minTCompressor))
            {
                MessageBox.Show(this,
                    "Минимальная Tcompr введена некорректно.",
                    "Параметры отчёта",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            ResultSettings.LabAssistantFullName = TbLabAssistant.Text.Trim();
            ResultSettings.TestName = TbTestName.Text.Trim();
            ResultSettings.MinPowerHighlight = minPower;
            ResultSettings.MinTCompressorHighlight = minTCompressor;

            DialogResult = true;
            Close();
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
