using ClosedXML.Excel;
using FridgeLabReport.Data;
using System.IO;

namespace FridgeLabReport
{
    internal static class Generator
    {
        private const int Line0 = 25;
        private const int Row0 = 2;

        private const int LineMinMaxAverage = 3;
        private const int RowStartMinMaxAverage = 5;
        private const int RowMin = 7;
        private const int RowMax = 8;
        private const int RowAverage = 9;


        public static void GenerateXlsx(
            string path,
            int Tcount,
            Dictionary<DataContainer.DataField, string> fields,
            List<DataContainer.DataRow> dataRows,
            ReportSettings? settings = null)
        {
            //settings ??= new ReportSettings();

            using var wb = new XLWorkbook(Path.Combine(AppContext.BaseDirectory, "Templates", $"t{Tcount}.xlsx"));
            IXLWorksheet ws = wb.Worksheet(1);

            ApplyWorkbookMetadata(ws, settings);


            bool isPowerMin = false;
            for (int line = 0; line < dataRows.Count; line++)
            {
                DataContainer.DataRow data = dataRows[line];
                int row = 0;
                int rowStart = 0;
                int rowEnd = 0;

                setCell(ws, line, ref row, toDateSec(data.LocalTime, "HH:mm:ss"));
                setCell(ws, line, ref row, toDate(data.Time, "dd.MM.yyyy"));
                setCell(ws, line, ref row, toDate(data.Time, "HH:mm:ss"));
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Pc]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Pe]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TcFilter]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TeSuction]]);


                double tCompressor = data[fields[DataContainer.DataField.TCompressor]];
                if(settings != null && settings.MinTCompressorHighlight.HasValue && tCompressor < settings.MinTCompressorHighlight.Value)
                {
                    setCell(ws, line, ref row, tCompressor, XLColor.DarkRed);
                }
                else
                {
                    setCell(ws, line, ref row, tCompressor, XLColor.FromHtml("#F4B183"));
                }

                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TCondInAir]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TCondOutAir]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TEvapInAir]], XLColor.FromHtml("#70AD47"));
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TEvapOutAir]], XLColor.FromHtml("#70AD47"));

                rowStart = Row0 + row;
                for (int i = 1; i <= Tcount; i++)
                {
                    XLColor color = XLColor.FromHtml("#C2C8FF");
                    if (i > 5) color = XLColor.FromHtml("#CCFFC2");
                    if (i > 10) color = XLColor.FromHtml("#FFFBC2");
                    if (i > 15) color = XLColor.FromHtml("#FFC2F3");
                    if (i > 20) color = XLColor.FromHtml("#C2F0FF");
                    if (i > 25) color = XLColor.FromHtml("#FFC2C2");
                    setCell(ws, line, ref row, data[fields[(DataContainer.DataField)i]], color);
                }
                rowEnd = Row0 + row - 1;
                string rangeT = $"{XLHelper.GetColumnLetterFromNumber(rowStart)}{Line0 + line}:{XLHelper.GetColumnLetterFromNumber(rowEnd)}{Line0 + line}";

                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Voltage]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Current]]);

                double power = data[fields[DataContainer.DataField.Power]];


                bool setYellow = false;
                int setYellowLineStart = 0;
                if (settings != null && settings.MinPowerHighlight.HasValue && power < settings.MinPowerHighlight.Value)
                {
                    setCell(ws, line, ref row, power, XLColor.Yellow);
                    if (!isPowerMin)
                    {
                        if(line > 0)
                        {
                            setYellow = true;
                            setYellowLineStart = Line0 + line;
                        }
                    }
                    isPowerMin = true;
                }
                else
                {
                    setCell(ws, line, ref row, power);
                    if (isPowerMin)
                    {
                        if(line < dataRows.Count - 1)
                        {
                            setYellow = true;
                            setYellowLineStart = Line0 + line - 1;
                        }
                    }
                    isPowerMin = false;
                }

                setCell(ws, line, ref row, data[fields[DataContainer.DataField.ChamberHumidity]], XLColor.FromHtml("#BFBFBF"));
                int r0 = Row0 + row;
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.ChamberTemperature]], XLColor.FromHtml("#BFBFBF"));

                setCellFormule(ws, line, ref row, $"MIN({rangeT})", XLColor.FromHtml("#DEEBF7"));
                setCellFormule(ws, line, ref row, $"AVERAGE({rangeT})", XLColor.FromHtml("#DEEBF7"));
                setCellFormule(ws, line, ref row, $"MAX({rangeT})", XLColor.FromHtml("#DEEBF7"));

                int l = Line0 + line;
                int r1 = Row0 + row;
                setCellFormule(ws, line, ref row, $"-42.5094 + 22.9586 * LN(E{l}) + 2.066199 * LN(E{l}) ^ 2 + 0.462774 * LN(E{l}) ^ 3");
                int r2 = Row0 + row;
                setCellFormule(ws, line, ref row, $"-42.5094 + 22.9586 * LN(F{l}) + 2.066199 * LN(F{l}) ^ 2 + 0.462774 * LN(F{l}) ^ 3");
                setCellFormule(ws, line, ref row, $"{XLHelper.GetColumnLetterFromNumber(r1)}{l}-G{l}");
                setCellFormule(ws, line, ref row, $"H{l}-{XLHelper.GetColumnLetterFromNumber(r2)}{l}");
                setCellFormule(ws, line, ref row, $"{XLHelper.GetColumnLetterFromNumber(r1)}{l}-{XLHelper.GetColumnLetterFromNumber(r0)}{l}");




                if (setYellow)
                {
                    SetYellowLine(ws, setYellowLineStart, Row0 + 3, Row0 + row - 1);
                }
            }


            for (int i = 0; i < 9; i++)
            {
                int row = RowStartMinMaxAverage + i;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }

            for (int i = 0; i < 5; i++)
            {
                int row = RowStartMinMaxAverage + i + 9 + Tcount;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i + 9, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i + 9, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i + 9, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }

            for (int i = 0; i < 5; i++)
            {
                int row = RowStartMinMaxAverage + i + 9 + 5 + 3 + Tcount;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }





            wb.SaveAs(path);
        }

        private static void ApplyWorkbookMetadata(IXLWorksheet ws, ReportSettings? settings)
        {
            if(settings != null)
            {
                ws.Cell(4, 15).Value = settings.TestName;
                ws.Cell(5, 15).Value = settings.LabAssistantFullName;
            }
        }


        private static void SetYellowLine(IXLWorksheet ws, int line, int start, int end)
        {
            for(int i = start; i <= end; i++)
            {
                IXLCell cell = ws.Cell(line, i);
                cell.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            }
        }




        private static bool TryGetFieldValue(
            DataContainer.DataRow data,
            Dictionary<DataContainer.DataField, string> fields,
            DataContainer.DataField field,
            out double value)
        {
            value = default;

            if (!fields.TryGetValue(field, out string? channelName))
                return false;

            value = data[channelName];
            return true;
        }

        private static void setCell(IXLWorksheet ws, int line, ref int row, double data, XLColor? color = null)
        {
            IXLCell cell = ws.Cell(Line0 + line, Row0 + row);

            cell.Value = Math.Round(data, 2);

            cell.Style.Border.SetLeftBorder(XLBorderStyleValues.Thick);
            cell.Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            if (color != null)
            {
                cell.Style.Fill.SetBackgroundColor(color);
            }

            row++;
        }

        private static void setCell(IXLWorksheet ws, int line, ref int row, long data, XLColor? color = null)
        {
            IXLCell cell = ws.Cell(Line0 + line, Row0 + row);

            cell.Value = data;

            cell.Style.Border.SetLeftBorder(XLBorderStyleValues.Thick);
            cell.Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            if (color != null)
            {
                cell.Style.Fill.SetBackgroundColor(color);
            }

            row++;
        }

        private static void setCell(IXLWorksheet ws, int line, ref int row, string data, XLColor? color = null)
        {
            IXLCell cell = ws.Cell(Line0 + line, Row0 + row);

            cell.Value = data;

            cell.Style.Border.SetLeftBorder(XLBorderStyleValues.Thick);
            cell.Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            if (color != null)
            {
                cell.Style.Fill.SetBackgroundColor(color);
            }

            row++;
        }

        private static void setCellFormule(IXLWorksheet ws, int line, ref int row, string formula, XLColor? color = null)
        {
            IXLCell cell = ws.Cell(Line0 + line, Row0 + row);

            cell.FormulaA1 = formula;

            cell.Style.Border.SetLeftBorder(XLBorderStyleValues.Thick);
            cell.Style.Border.SetRightBorder(XLBorderStyleValues.Thick);
            if (color != null)
            {
                cell.Style.Fill.SetBackgroundColor(color);
            }

            row++;
        }

        private static string toDateSec(long time, string format)
        {
            DateTime dt = new DateTime().AddMilliseconds(time);
            return dt.ToString(format);
        }

        private static string toDate(long time, string format)
        {
            DateTime dt = DateTimeOffset.FromUnixTimeMilliseconds(time).LocalDateTime;
            return dt.ToString(format);
        }
    }
}
