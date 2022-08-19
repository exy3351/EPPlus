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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// 
    /// </summary>
    public class DoubleEnumerableArgConverter : CollectionFlattener<ExcelDoubleCellValue>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ignoreHidden"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        public virtual IEnumerable<ExcelDoubleCellValue> ConvertArgs(bool ignoreHidden, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreNonNumeric = false)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
                {
                    if (arg.IsExcelRange)
                    {
                        foreach (var cell in arg.ValueAsRangeInfo)
                        {
                            if(!ignoreErrors && cell.IsExcelError) throw new ExcelErrorValueException(ExcelErrorValue.Parse(cell.Value.ToString()));
                            if (!CellStateHelper.ShouldIgnore(ignoreHidden, ignoreNonNumeric, cell, context) && ConvertUtil.IsNumericOrDate(cell.Value))
                            {
                                var val = new ExcelDoubleCellValue(cell.ValueDouble, cell.Row, cell.Column);
                                argList.Add(val);
                            }       
                        }
                    }
                    else
                    {
                        if(!ignoreErrors && arg.ValueIsExcelError) throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
                        if (ConvertUtil.IsNumericOrDate(arg.Value) && !CellStateHelper.ShouldIgnore(ignoreHidden, arg, context))
                        {
                            var val = new ExcelDoubleCellValue(ConvertUtil.GetValueDouble(arg.Value));
                            argList.Add(val);
                        }
                    }
                });
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="ignoreHidden"></param>
        /// <returns></returns>
        public virtual IEnumerable<ExcelDoubleCellValue> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments, bool ignoreHidden)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                //var cellInfo = arg.Value as EpplusExcelDataProvider.CellInfo;
                //var value = cellInfo != null ? cellInfo.Value : arg.Value;
                if (arg.Value is IRangeInfo)
                {
                    foreach (var cell in (IRangeInfo)arg.Value)
                    {
                        if((!ignoreHidden && cell.IsHiddenRow) || !cell.IsHiddenRow)
                        {
                            var val = new ExcelDoubleCellValue(cell.ValueDoubleLogical, cell.Row, cell.Column);
                            argList.Add(val);
                        }
                        
                    }
                }
                else
                {
                    if (arg.Value is double || arg.Value is int || arg.Value is bool)
                    {
                        argList.Add(Convert.ToDouble(arg.Value));
                    }
                    else if (arg.Value is string)
                    {
                        argList.Add(0d);
                    }
                }
            });
        }
    }
}
