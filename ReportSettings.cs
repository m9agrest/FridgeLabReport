namespace FridgeLabReport
{
    public sealed class ReportSettings
    {
        public string LabAssistantFullName { get; set; } = string.Empty;
        public string TestName { get; set; } = string.Empty;
        public double? MinPowerHighlight { get; set; }
        public double? MinTCompressorHighlight { get; set; }
        public double? MinAllT { get; set; }

        public ReportSettings Clone()
        {
            return new ReportSettings
            {
                LabAssistantFullName = LabAssistantFullName,
                TestName = TestName,
                MinPowerHighlight = MinPowerHighlight,
                MinTCompressorHighlight = MinTCompressorHighlight,
                MinAllT = MinAllT,
            };
        }

        public bool IsEmpty()
        {
            return string.IsNullOrWhiteSpace(LabAssistantFullName)
                && string.IsNullOrWhiteSpace(TestName)
                && MinPowerHighlight == null
                && MinTCompressorHighlight == null
                && MinAllT == null;
        }
    }
}
