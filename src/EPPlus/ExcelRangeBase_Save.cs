/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Data;
using OfficeOpenXml.Export.ToDataTable;
using System.Threading.Tasks;

namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the <see cref="ToDataTableOptions.Default"/> settings.
        /// </summary>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable()
        {
            return ToDataTable(ToDataTableOptions.Default);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="configHandler">Configures the settings used to convert the range.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler)
        {
            var o = ToDataTableOptions.Default;
            configHandler.Invoke(o);
            return ToDataTable(o);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="options">Sets the settings used to convert the range.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(ToDataTableOptions options)
        {
            var func = new ToDataTable(options, this);
            return func.Execute();
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="configHandler">Configures the settings used to convert the range.</param>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(Action<ToDataTableOptions> configHandler, DataTable dataTable)
        {
            var o = ToDataTableOptions.Default;
            configHandler.Invoke(o);
            return ToDataTable(o, dataTable);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(DataTable dataTable)
        {
            return ToDataTable(ToDataTableOptions.Default, dataTable);
        }

        /// <summary>
        /// Returns the range as a <see cref="DataTable"/> with the option supplied.
        /// </summary>
        /// <param name="options">Sets the settings used to convert the range.</param>
        /// <param name="dataTable">The data table to add the range data to.</param>
        /// <returns>A <see cref="DataTable"/> representing the range.</returns>
        public DataTable ToDataTable(ToDataTableOptions options, DataTable dataTable)
        {
            var func = new ToDataTable(options, this);
            return func.Execute(dataTable);
        }

        /// <summary>
        /// Converts a range to text in CSV format.
        /// </summary>
        /// <returns>A string containing the text</returns>
        public string ToText()
        {
            return ToText(null);
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the csv text</param>
        /// <returns>A string containing the text</returns>
        public string ToText(ExcelOutputTextFormat Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                SaveToText(ms, Format);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return sr.ReadToEnd();
            }
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public void SaveToText(FileInfo file, ExcelOutputTextFormat Format)
        {
            var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write);
            SaveToText(fileStream, Format);
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public void SaveToText(Stream stream, ExcelOutputTextFormat Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormat();
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) sw.Write(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;

            bool hasTextQ = Format.TextQualifier != '\0';
            string doubleTextQualifiers = new string(Format.TextQualifier, 2);
            var skipLinesBegining = Format.SkipLinesBeginning+(Format.FirstRowIsHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = _fromRow; row <= _toRow; row++)
            {
                if (row == _fromRow && Format.FirstRowIsHeader)
                {
                    sw.Write(WriteHeaderRow(Format, hasTextQ, row, ci));
                    continue;
                }


                if (SkipLines(Format, row, skipLinesBegining))
                {
                    continue;
                }

                for (int col = _fromCol; col <= _toCol; col++)
                {
                    string t = GetText(Format, maxFormats, ci, row, col, out bool isText);

                    if (hasTextQ && isText)
                    {
                        sw.Write(Format.TextQualifier);
                        sw.Write(t.Replace(Format.TextQualifier.ToString(), doubleTextQualifiers));
                        sw.Write(Format.TextQualifier);
                    }
                    else
                    {
                        sw.Write(t);
                    }
                    if (col != _toCol) sw.Write(Format.Delimiter);
                }
                if (row != _toRow-Format.SkipLinesEnd) sw.Write(Format.EOL);
            }
            if (!string.IsNullOrEmpty(Format.Footer)) sw.Write(Format.EOL + Format.Footer);
            sw.Flush();
        }

        /// <summary>
        /// Converts a range to text in CSV format.
        /// </summary>
        /// <returns>A string containing the text</returns>
        public async Task<string> ToTextAsync()
        {
            return await ToTextAsync(null).ConfigureAwait(false);
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="Format">Information how to create the csv text</param>
        /// <returns>A string containing the text</returns>
        public async Task<string> ToTextAsync(ExcelOutputTextFormat Format)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await SaveToTextAsync(ms, Format).ConfigureAwait(false);
                ms.Position = 0;
                var sr = new StreamReader(ms);
                return await sr.ReadToEndAsync().ConfigureAwait(false);
            }
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="file">The file to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public async Task SaveToTextAsync(FileInfo file, ExcelOutputTextFormat Format)
        {
            var fileStream = file.Open(FileMode.Create, FileAccess.Write, FileShare.Write);
            await SaveToTextAsync(fileStream, Format).ConfigureAwait(false);
        }
        /// <summary>
        /// Converts a range to text in CSV format.
        /// Invariant culture is used by default.
        /// </summary>
        /// <param name="stream">The strem to write to</param>
        /// <param name="Format">Information how to create the csv text</param>
        public async Task SaveToTextAsync(Stream stream, ExcelOutputTextFormat Format)
        {
            if (Format == null) Format = new ExcelOutputTextFormat();
            var sw = new StreamWriter(stream, Format.Encoding);
            if (!string.IsNullOrEmpty(Format.Header)) sw.Write(Format.Header + Format.EOL);
            int maxFormats = Format.Formats == null ? 0 : Format.Formats.Length;

            bool hasTextQ = Format.TextQualifier != '\0';
            string encodedTextQualifier="";
            if (hasTextQ)
            {
                if (Format.EncodedTextQualifiers == null)
                {
                    encodedTextQualifier = new string(Format.TextQualifier, 2);
                }
                else
                {
                    encodedTextQualifier = Format.EncodedTextQualifiers;
                }
            }
            var skipLinesBegining = Format.SkipLinesBeginning + (Format.FirstRowIsHeader ? 1 : 0);
            CultureInfo ci = GetCultureInfo(Format);
            for (int row = _fromRow; row <= _toRow; row++)
            {
                if (row == _fromRow && Format.FirstRowIsHeader)
                {
                    await sw.WriteAsync(WriteHeaderRow(Format, hasTextQ, row, ci)).ConfigureAwait(false);
                    continue;
                }

                if (SkipLines(Format, row, skipLinesBegining))
                {
                    continue;
                }

                for (int col = _fromCol; col <= _toCol; col++)
                {
                    string t = GetText(Format, maxFormats, ci, row, col, out bool isText);

                    if (hasTextQ && isText)
                    {
                        await sw.WriteAsync(Format.TextQualifier + t.Replace(Format.TextQualifier.ToString(), encodedTextQualifier) + Format.TextQualifier).ConfigureAwait(false);
                    }
                    else
                    {
                        await sw.WriteAsync(t).ConfigureAwait(false);
                    }
                    if (col != _toCol) await sw.WriteAsync(Format.Delimiter).ConfigureAwait(false); 
                }
                if (row != _toRow - Format.SkipLinesEnd) await sw.WriteAsync(Format.EOL).ConfigureAwait(false); 
            }
            if (!string.IsNullOrEmpty(Format.Footer)) await sw.WriteAsync(Format.EOL + Format.Footer).ConfigureAwait(false); 
            sw.Flush();
        }

        /// <summary>
        /// Returns the range as JSON
        /// </summary>
        /// <returns>A JSON string</returns>
        public string ToJson()
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            var ms = RecyclableMemory.GetStream();
            re.Export(ms);
            return Encoding.UTF8.GetString(ms.ToArray());
        }
        /// <summary>
        /// Returns the range as JSON
        /// </summary>
        /// <param name="settings">Configures settings for the JSON export</param>
        /// <returns></returns>
        public string ToJson(Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            var ms = RecyclableMemory.GetStream();
            re.Export(ms);
            return s.Encoding.GetString(ms.ToArray());
        }
        /// <summary>
        /// Saves the range as JSON to a stream.
        /// </summary>
        /// <param name="stream">The writable stream to write the JSON to.</param>
        public void SaveToJson(Stream stream)
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            re.Export(stream);
        }
        /// <summary>
        /// Saves the range as JSON to a stream.
        /// </summary>
        /// <param name="stream">The writable stream to write the JSON to</param>
        /// <param name="settings">Configures settings for the JSON export</param>
        public void SaveToJson(Stream stream, Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            re.Export(stream);
        }

        /// <summary>
        /// Save the range to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream)
        {
            var re = new JsonRangeExport(this, new JsonRangeExportSettings());
            await re.ExportAsync(stream);
        }
        /// <summary>
        /// Save the range to json
        /// </summary>
        /// <param name="stream">The stream to save to.</param>
        /// <param name="settings">Settings for the json output.</param>
        /// <returns></returns>
        public async Task SaveToJsonAsync(Stream stream, Action<JsonRangeExportSettings> settings)
        {
            var s = new JsonRangeExportSettings();
            settings.Invoke(s);
            var re = new JsonRangeExport(this, s);
            await re.ExportAsync(stream);
        }

        private static CultureInfo GetCultureInfo(ExcelOutputTextFormat Format)
        {
            var ci = (CultureInfo)(Format.Culture.Clone() ?? CultureInfo.InvariantCulture.Clone());
            if (Format.DecimalSeparator != null)
            {
                ci.NumberFormat.NumberDecimalSeparator = Format.DecimalSeparator;
            }
            if (Format.ThousandsSeparator != null)
            {
                ci.NumberFormat.NumberGroupSeparator = Format.ThousandsSeparator;
            }

            return ci;
        }

        private bool SkipLines(ExcelOutputTextFormat Format, int row, int skipLinesBegining)
        {
            return skipLinesBegining > row - _fromRow ||
                               Format.SkipLinesEnd > _toRow - row;
        }

        private string GetText(ExcelOutputTextFormat Format, int maxFormats, CultureInfo ci, int row, int col, out bool isText)
        {
            var v=GetCellStoreValue(row, col);

            var ix = col - _fromCol;
            isText = false;
            string fmt;
            if (ix < maxFormats)
            {
                fmt = Format.Formats[ix];
            }
            else
            {
                fmt = "";
            }
            string t;

            if (string.IsNullOrEmpty(fmt))
            {
                if (Format.UseCellFormat)
                {
                    t = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);
                    if (!ConvertUtil.IsNumericOrDate(v._value)) isText = true;
                }
                else
                {
                    if (ConvertUtil.IsNumeric(v._value))
                    {
                        t = ConvertUtil.GetValueDouble(v._value).ToString("r", ci);
                    }
                    else if (v._value is DateTime date)
                    {
                        t = date.ToString("G", ci);
                    }
                    else
                    {
                        t = v._value.ToString();
                        isText = true;
                    }
                }
            }
            else
            {
                if (fmt == "$")
                {
                    if (Format.UseCellFormat)
                    {
                        t = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);
                    }
                    else
                    {
                        t = v._value.ToString();
                    }
                    isText = true;
                }
                else if (ConvertUtil.IsNumeric(v._value))
                {
                    t = ConvertUtil.GetValueDouble(v._value).ToString(fmt, ci);
                }
                else if (v._value is DateTime date)
                {
                    t = date.ToString(fmt, ci);
                }
                else if (v._value is TimeSpan ts)
                {
                    t = ts.ToString(fmt, ci);
                }
                else
                {
                    t = v._value.ToString();
                    isText = true;
                }
            }

            //If a formatted numeric/date value contains the delimitter or a text qualifier treat it as text.
            if (isText == false && string.IsNullOrEmpty(t)==false && t.IndexOfAny(new []{ Format.Delimiter, Format.TextQualifier}) >= 0)
            {
                isText = true;
            }
            return t;
        }

        private Core.CellStore.ExcelValue GetCellStoreValue(int row, int col)
        {
            var v = _worksheet.GetCoreValueInner(row, col);
            if (_worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
            {
                var xml = new XmlDocument();
                XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + v._value.ToString() + "</d:si>", Encoding.UTF8);
                var rt = new ExcelRichTextCollection(_worksheet.NameSpaceManager, xml.SelectSingleNode("d:si", _worksheet.NameSpaceManager), this);
                v._value = rt.Text;
            }
            return v;
        }

        private string WriteHeaderRow(ExcelOutputTextFormat Format, bool hasTextQ, int row, CultureInfo ci)
        {
            var sb = new StringBuilder();
            for (int col = _fromCol; col <= _toCol; col++)
            {
                var v = GetCellStoreValue(row, col);
                var s = ValueToTextHandler.GetFormattedText(v._value, _workbook, v._styleId, false, ci);

                if (hasTextQ)
                {
                    sb.Append(Format.TextQualifier);
                    sb.Append(s.Replace(Format.TextQualifier.ToString(), new string(Format.TextQualifier, 2)));
                    sb.Append(Format.TextQualifier);
                }
                else
                {
                    sb.Append(s);
                }

                if(col < _toCol)
                    sb.Append(Format.Delimiter);
            }
            if (row != _toRow) sb.Append(Format.EOL);
            return sb.ToString();
        }
    }
}