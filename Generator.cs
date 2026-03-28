using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using FridgeLabReport.Data;
using System.Globalization;
using System.IO;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using WTableBorders = DocumentFormat.OpenXml.Wordprocessing.TableBorders;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

using WTopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using WBottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using WLeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using WRightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using WInsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using WInsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;

using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using WDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using WBreak = DocumentFormat.OpenXml.Wordprocessing.Break;
using WTabChar = DocumentFormat.OpenXml.Wordprocessing.TabChar;

namespace FridgeLabReport
{
    internal static class Generator
    {













        // ================================================================================================== //
        // ===================================ГЕНЕРАТОР XLSX ФАЙЛА=========================================== //
        // ================================================================================================== //





        private const int Line0 = 25;
        private const int Row0 = 2;

        private const int LineMinMaxAverage = 3;//линия для колонки подсчетов минимума максима и среднего
        private const int RowStartMinMaxAverage = 5;//с какой колонки будет заполнять // основная таблица, начало полезных данных
        private const int RowMin = 7;//колонка для подсчета минимума
        private const int RowMax = 8;//колонка для подсчета максимума
        private const int RowAverage = 9;//колонка для подсчета среднего


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

            ApplyWorkbookMetadata(ws, settings, dataRows);

            /*********    Генерируем основную таблицу     *********/

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
                if (settings != null && settings.MinTCompressorHighlight.HasValue && tCompressor < settings.MinTCompressorHighlight.Value)//TCompressor ниже минимума
                {
                    setCell(ws, line, ref row, tCompressor, XLColor.DarkRed);
                }
                else//TCompressor выше минимума
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
                    //каждые 5 колонок у Т, меняем цвет
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


                bool setYellow = false;//нужно ли рисовать линию
                int setYellowLineStart = 0;

                if (settings != null && settings.MinPowerHighlight.HasValue && power < settings.MinPowerHighlight.Value)//мощность ниже минимума
                {
                    setCell(ws, line, ref row, power, XLColor.Yellow);
                    if (!isPowerMin)
                    {
                        if (line > 0)//если не первая строка, то нужно рисовать линию
                        {
                            setYellow = true;
                            setYellowLineStart = Line0 + line;
                        }
                    }
                    isPowerMin = true;
                }
                else//мощность выше минимума
                {
                    setCell(ws, line, ref row, power);
                    if (isPowerMin)
                    {
                        if (line < dataRows.Count - 1)//если не последняя строка, то нужно рисовать линию
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

                //мин макс и сред значение по линии от Т1 до Tcount
                setCellFormule(ws, line, ref row, $"MIN({rangeT})", XLColor.FromHtml("#DEEBF7"));
                setCellFormule(ws, line, ref row, $"AVERAGE({rangeT})", XLColor.FromHtml("#DEEBF7"));
                setCellFormule(ws, line, ref row, $"MAX({rangeT})", XLColor.FromHtml("#DEEBF7"));

                //генерируем формулы последних 5 колонок
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

            /********   Генерируем таблицу мин макс сред значений по основной таблице    *******/

            //заполняем записи, до Т1
            for (int i = 0; i < 9; i++)
            {
                int row = RowStartMinMaxAverage + i;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }

            //пропускаем от Т1 до Tcount

            //пишем 5 значений до 3 колонок мин мак сред в строках Т1 - Tcount
            for (int i = 0; i < 5; i++)
            {
                int row = RowStartMinMaxAverage + i + 9 + Tcount;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i + 9, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i + 9, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i + 9, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }

            //пропускам 3 колонки мин мак сред в строках Т1 - Tcount

            //пишем 5 полследних колонок
            for (int i = 0; i < 5; i++)
            {
                int row = RowStartMinMaxAverage + i + 9 + 5 + 3 + Tcount;
                string columns = $"{XLHelper.GetColumnLetterFromNumber(row)}{Line0}:{XLHelper.GetColumnLetterFromNumber(row)}{Line0 + dataRows.Count - 1}";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowMin).FormulaA1 = $"MIN({columns})";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowMax).FormulaA1 = $"MAX({columns})";
                ws.Cell(LineMinMaxAverage + i + 9 + 5, RowAverage).FormulaA1 = $"AVERAGE({columns})";
            }





            wb.SaveAs(path);


            //GenerateDocx(path + ".docx", Tcount, dataRows, ws, settings);
        }


        //отрисовывает желтую строку
        private static void ApplyWorkbookMetadata(IXLWorksheet ws, ReportSettings? settings, List<DataContainer.DataRow> dataRows)
        {
            if (settings != null)
            {
                ws.Cell(4, 15).Value = settings.TestName;
                ws.Cell(5, 15).Value = settings.LabAssistantFullName;
            }

            ws.Cell(7, 15).Value = toDate(dataRows[0].StartTime, "dd.MM.yyyy HH:mm:ss");
            ws.Cell(8, 15).Value = toDate(dataRows[0].Time, "dd.MM.yyyy HH:mm:ss") + " - " + toDate(dataRows.Last().Time, "dd.MM.yyyy HH:mm:ss");
            ws.Cell(9, 15).Value = toDateSec(dataRows.Last().Time - dataRows[0].Time, "HH:mm:ss");

        }


        private static void SetYellowLine(IXLWorksheet ws, int line, int start, int end)
        {
            for (int i = start; i <= end; i++)
            {
                IXLCell cell = ws.Cell(line, i);
                cell.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            }
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









        // ================================================================================================== //
        // ===================================ГЕНЕРАТОР DOCX ФАЙЛА=========================================== //
        // ================================================================================================== //


        static int DocLine1 => LineMinMaxAverage - 1;
        static int DocRowText => RowAverage - 7;
        static int DocLine2 => LineMinMaxAverage + 18;


        public static void GenerateDocx(string docxPath, int Tcount, List<DataContainer.DataRow> dataRows, IXLWorksheet ws, ReportSettings? settings = null)
        {
            // Пересчитать формулы
            ws.Workbook.RecalculateAllFormulas();

            using var doc = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();

            if (mainPart.Document == null)
                mainPart.Document = new Document(new Body());

            Body body = mainPart.Document.Body!;


            // ------ Заголовок ------
            var title = new Paragraph(new WRun(new WText("Отчет"), new WBreak(), new WBreak()))
            {
                ParagraphProperties = new ParagraphProperties
                {
                    Justification = new Justification { Val = JustificationValues.Center }
                }
            };
            body.Append(title);


            // ------ Информация ------
            if (settings != null)
            {
                Paragraph? baseInfo = null;
                if (settings.TestName != null)
                {
                    baseInfo = new Paragraph();
                    baseInfo.Append(new WRun(new WText("Название испытания: " + settings.TestName), new WBreak()));
                }
                if (settings.LabAssistantFullName != null)
                {
                    if (baseInfo == null) baseInfo = new Paragraph();
                    baseInfo.Append(new WRun(new WText("Лаборант: " + settings.LabAssistantFullName), new WBreak()));
                }
                if (baseInfo != null)
                {
                    baseInfo.Append(new WRun(new WBreak()));
                    body.Append(baseInfo);
                }
            }


            // ------ Время испытания ------
            var timeInfo = new Paragraph();
            timeInfo.Append(new WRun(new WText("Начало испытания: " + toDate(dataRows[0].StartTime, "dd.MM.yyyy HH:mm:ss")), new WBreak()));
            timeInfo.Append(new WRun(new WText("Выбранный отрезок:"), new WBreak()));
            timeInfo.Append(new WRun(new WTabChar(), new WText("с: " + toDate(dataRows[0].Time, "dd.MM.yyyy HH:mm:ss")), new WBreak()));
            timeInfo.Append(new WRun(new WTabChar(), new WText("по: " + toDate(dataRows.Last().Time, "dd.MM.yyyy HH:mm:ss")), new WBreak()));
            timeInfo.Append(new WRun(new WTabChar(), new WText("итого: " + toDateSec(dataRows.Last().LocalTime - dataRows[0].LocalTime, "HH:mm:ss")), new WBreak(), new WBreak()));
            body.Append(timeInfo);


            // ------ Таблица ------
            var tableInfo = new Paragraph();
            var table = MakeTable(ws);
            var tableEnd = new Paragraph();
            tableInfo.Append(new WRun(new WText("Минимальные и максимальные значения с датчиков:")));
            tableEnd.Append(new WRun(new WBreak(), new WBreak()));
            body.Append(tableInfo);
            body.Append(table);
            body.Append(tableEnd);



            // ------ min TCompr ------
            if (settings != null && settings.MinTCompressorHighlight.HasValue)
            {
                var minTComprInfo = new Paragraph();
                minTComprInfo.Append(new WRun(new WText($"Падения TCompressor, ниже минимального: {settings.MinTCompressorHighlight.Value}"), new WBreak()));
                string? start = null;
                string? end = null;
                bool isStart = false;
                for (int  i = 0; i < dataRows.Count; i++)
                {
                    var data = dataRows[i];
                    double tCompressor = data[DataContainer.DataField.TCompressor];
                    if (tCompressor < settings.MinTCompressorHighlight.Value)
                    {
                        end = toDate(data.Time, "dd.MM.yyyy HH:mm:ss");
                        if (i == 0)
                        {
                            isStart = true;
                        }
                        if(start == null)
                        {
                            start = end;
                        }
                    }
                    else
                    {
                        if (start != null)
                        {
                            if (isStart)
                            {
                                isStart = false;

                                minTComprInfo.Append(new WRun(new WTabChar(), new WText("с начала выбранного отрезка, по: " + end), new WBreak()));
                            }
                            else
                            {
                                minTComprInfo.Append(new WRun(new WTabChar(), new WText("с: " + start + "; по: " + end), new WBreak()));
                            }
                            start = null;
                            end = null;
                        }
                    }
                }
                if(start != null)
                {
                    if (isStart)
                    {
                        minTComprInfo.Append(new WRun(new WTabChar(), new WText("На протяжении всего выбранного отрезка"), new WBreak()));
                    }
                    else
                    {
                        minTComprInfo.Append(new WRun(new WTabChar(), new WText("с: " + start + "; по конец выбранного отрезка"), new WBreak()));
                    }
                }

                body.Append(minTComprInfo);
            }

            // ------ min Power ------
            if (settings != null && settings.MinPowerHighlight.HasValue)
            {
                var minPower = new Paragraph();
                minPower.Append(new WRun(new WText($"Падения мощности, ниже минимального: {settings.MinPowerHighlight.Value}"), new WBreak()));
                string? start = null;
                string? end = null;
                bool isStart = false;
                for (int i = 0; i < dataRows.Count; i++)
                {
                    var data = dataRows[i];
                    double power = data[DataContainer.DataField.Power];
                    if (power < settings.MinPowerHighlight.Value)
                    {
                        end = toDate(data.Time, "dd.MM.yyyy HH:mm:ss");
                        if (i == 0)
                        {
                            isStart = true;
                        }
                        if (start == null)
                        {
                            start = end;
                        }
                    }
                    else
                    {
                        if (start != null)
                        {
                            if (isStart)
                            {
                                isStart = false;

                                minPower.Append(new WRun(new WTabChar(), new WText("с начала выбранного отрезка, по: " + end), new WBreak()));
                            }
                            else
                            {
                                minPower.Append(new WRun(new WTabChar(), new WText("с: " + start + "; по: " + end), new WBreak()));
                            }
                            start = null;
                            end = null;
                        }
                    }
                }
                if (start != null)
                {
                    if (isStart)
                    {
                        minPower.Append(new WRun(new WTabChar(), new WText("На протяжении всего выбранного отрезка"), new WBreak()));
                    }
                    else
                    {
                        minPower.Append(new WRun(new WTabChar(), new WText("с: " + start + "; по конец выбранного отрезка"), new WBreak()));
                    }
                }

                body.Append(minPower);
            }



            mainPart.Document.Save();
        }


        private static WTable MakeTable(IXLWorksheet ws)
        {
            var table = new WTable();

            // Границы таблицы
            var props = new TableProperties(
                new TableBorders(
                    new WTopBorder { Val = BorderValues.Single, Size = 4 },
                    new WBottomBorder { Val = BorderValues.Single, Size = 4 },
                    new WLeftBorder { Val = BorderValues.Single, Size = 4 },
                    new WRightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            );

            table.AppendChild(props);

            var culture = new CultureInfo("ru-RU");

            for (int r = DocLine1; r <= DocLine2; r++)
            {
                var tr = new TableRow();

                tr.Append(MakeCell(ws.Cell(r, DocRowText).GetFormattedString(culture), "5000"));
                tr.Append(MakeCell(ws.Cell(r, RowMin).GetFormattedString(culture), "1200"));
                tr.Append(MakeCell(ws.Cell(r, RowMax).GetFormattedString(culture), "1200"));
                tr.Append(MakeCell(ws.Cell(r, RowAverage).GetFormattedString(culture), "1200"));

                table.Append(tr);
            }



            return table;
        }

        private static TableCell MakeCell(string text, string width)
        {
            return new TableCell(
                new TableCellProperties(
                    new TableCellWidth
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = width
                    }
                ),
                new Paragraph(
                    new WRun(
                        new WText(text ?? string.Empty)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        }
                    )
                )
            );
        }

    }
}
