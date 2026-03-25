using ClosedXML.Excel;
using ClosedXML.Parser;
using FridgeLabReport.Data;
using System.IO;
using System.Windows.Controls;

namespace FridgeLabReport
{
    internal static class Generator
    {
        const int Line0 = 24;
        const int Row0 = 2;

        public static void GenerateXlsx(string path, int Tcount, Dictionary<DataContainer.DataField, string> fields, List<DataContainer.DataRow> dataRows)
        {
            
            using var wb = new XLWorkbook(Path.Combine(AppContext.BaseDirectory, "Templates", $"t{Tcount}.xlsx"));
            IXLWorksheet ws = wb.Worksheet(1);

            //читаем файл 't' + Tcount + ".xlsx";
            for (int line = 0; line < dataRows.Count; line++)
            {
                DataContainer.DataRow data = dataRows[line];
                int row = 0;
                int rowStart = 0;
                int rowEnd = 0;

                setCell(ws, line, ref row, toDate(data.LocalTime, "HH:mm:ss"));//hh:mm:ss
                setCell(ws, line, ref row, toDate(data.Time, "dd.MM.yyyy"));//dd:mm:yyyy
                setCell(ws, line, ref row, toDate(data.Time, "HH:mm:ss"));//hh:mm:ss
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Pc]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Pe]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TcFilter]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TeSuction]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TCompressor]], XLColor.FromHtml("#F4B183"));
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TCondInAir]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TCondOutAir]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TEvapInAir]], XLColor.FromHtml("#70AD47"));
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.TEvapOutAir]], XLColor.FromHtml("#70AD47"));

                rowStart = Row0 + row;
                for (int i = 1; i <= Tcount; i++)
                {
                    setCell(ws, line, ref row, data[fields[(DataContainer.DataField)i]]);
                }
                rowEnd = Row0 + row - 1;
                string rangeT = $"{XLHelper.GetColumnLetterFromNumber(rowStart)}{Line0 + line}:{XLHelper.GetColumnLetterFromNumber(rowEnd)}{Line0 + line}";

                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Voltage]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Current]]);
                setCell(ws, line, ref row, data[fields[DataContainer.DataField.Power]]);
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
            }

            wb.SaveAs(path);

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

        private static string toDate(long time, string format)
        {
            DateTime dt = DateTimeOffset.FromUnixTimeMilliseconds(time).LocalDateTime;
            return dt.ToString(format);
        }


    }
}
