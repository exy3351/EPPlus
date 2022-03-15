﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;
namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class ExcelHtmlRangeExporter : HtmlExporterBase
    {        
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms, 0);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync(int rangeIndex)
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderHtmlAsync(ms, rangeIndex);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write the css to.</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <exception cref="IOException"></exception>
        public async Task RenderHtmlAsync(Stream stream, int rangeIndex)
        {
            ValidateRangeIndex(rangeIndex);
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            _mergedCells.Clear();
            var range = _ranges[rangeIndex];
            GetDataTypes(_ranges[rangeIndex]);

            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);
            AddClassesAttributes(writer);
            AddTableAccessibilityAttributes(Settings, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral==eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await SetColumnGroupAsync(writer, range, Settings);
            }

            if (Settings.HeaderRows > 0 || Settings.Headers.Count > 0)
            {
                await RenderHeaderRowAsync(range, writer);
            }
            // table rows
            await RenderTableRowsAsync(range, writer);

            // end tag table
            await writer.RenderEndTagAsync();

        }

        /// <summary>
        /// Renders the first range of the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = await GetHtmlStringAsync();
            var css = await GetCssStringAsync();
            return string.Format(htmlDocument, html, css);
        }
        private async Task RenderTableRowsAsync(ExcelRangeBase range, EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Tbody);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var row = range._fromRow + Settings.HeaderRows;
            var endRow = range._toRow;
            var ws = range.Worksheet;
            HtmlImage image = null;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    writer.AddAttribute("scope", "row");
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix);
                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    SetColRowSpan(range, writer, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        await _cellDataWriter.WriteAsync(cell, dataType, writer, Settings, false, image);
                    }
                    else
                    {
                        await writer.RenderBeginTagAsync(HtmlElements.TableData);
                        var imageCellClassName = image == null ? "" : Settings.StyleClassPrefix + "image-cell";
                        writer.SetClassAttributeFromStyle(cell, false, Settings, imageCellClassName);
                        await RenderHyperlinkAsync(writer, cell);
                        await writer.RenderEndTagAsync();
                        await writer.ApplyFormatAsync(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(Settings.Minify);
                row++;
            }

            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            // end tag tbody
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }

        private async Task RenderHeaderRowAsync(ExcelRangeBase range, EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
            var headerRows = Settings.HeaderRows == 0 ? 1 : Settings.HeaderRows;
            HtmlImage image = null;
            for (int i = 0; i < headerRows; i++)
            {
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                }
                var row = range._fromRow + i;
                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix);
                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var cell = range.Worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        writer.AddAttribute("data-datatype", _datatypes[col - range._fromCol]);
                    }
                    SetColRowSpan(range, writer, cell);
                    if (Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = image == null ? "" : Settings.StyleClassPrefix + "image-cell";
                        writer.SetClassAttributeFromStyle(cell, true, Settings, imageCellClassName);
                    }
                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell._fromRow, cell._fromCol);
                    }
                    await AddImageAsync(writer, Settings, image, cell.Value);

                    await writer.RenderBeginTagAsync(HtmlElements.TableHeader);
                    if (Settings.HeaderRows > 0)
                    {
                        if (cell.Hyperlink == null)
                        {
                            await writer.WriteAsync(GetCellText(cell));
                        }
                        else
                        {
                            await RenderHyperlinkAsync(writer, cell);
                        }
                    }
                    else if (Settings.Headers.Count < col)
                    {
                        await writer.WriteAsync(Settings.Headers[col]);
                    }

                    await writer.RenderEndTagAsync();
                    await writer.ApplyFormatAsync(Settings.Minify);
                }
                writer.Indent--;
                await writer.RenderEndTagAsync();
            }
            await writer.ApplyFormatDecreaseIndentAsync(Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(Settings.Minify);
        }
        private async Task RenderHyperlinkAsync(EpplusHtmlWriter writer, ExcelRangeBase cell)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    writer.AddAttribute("href", eurl.AbsolutePath);
                    await writer.RenderBeginTagAsync(HtmlElements.A);
                    await writer.WriteAsync(eurl.Display);
                    await writer.RenderEndTagAsync();
                }
                else
                {
                    //Internal
                    await writer.WriteAsync(GetCellText(cell));
                }
            }
            else
            {
                writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                await writer.RenderBeginTagAsync(HtmlElements.A);
                await writer.WriteAsync(GetCellText(cell));
                await writer.RenderEndTagAsync();
            }
        }
    }
}
#endif

