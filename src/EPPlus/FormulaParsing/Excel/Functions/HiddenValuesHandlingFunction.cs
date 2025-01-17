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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for functions that needs to handle cells that is not visible.
    /// </summary>
    public abstract class HiddenValuesHandlingFunction : ExcelFunction
    {
        /// <summary>
        /// 
        /// </summary>
        public HiddenValuesHandlingFunction()
        {
            IgnoreErrors = true;
        }
        /// <summary>
        /// Set to true or false to indicate whether the function should ignore hidden values.
        /// </summary>
        public bool IgnoreHiddenValues
        {
            get;
            set;
        }

        /// <summary>
        /// Set to true to indicate whether the function should ignore error values
        /// </summary>
        public bool IgnoreErrors
        {
            get; set;
        }

        /// <inheritdoc/>
        protected override IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return ArgsToDoubleEnumerable(arguments, context, IgnoreErrors, false);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <param name="ignoreErrors"></param>
        /// <param name="ignoreNonNumeric"></param>
        /// <returns></returns>
        protected IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context, bool ignoreErrors, bool ignoreNonNumeric)
        {
            if (!arguments.Any())
            {
                return Enumerable.Empty<ExcelDoubleCellValue>();
            }
            if (IgnoreHiddenValues)
            {
                var nonHidden = arguments.Where(x => !x.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
                return base.ArgsToDoubleEnumerable(IgnoreHiddenValues, nonHidden, context);
            }
            return base.ArgsToDoubleEnumerable(IgnoreHiddenValues, ignoreErrors, arguments, context, ignoreNonNumeric);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="c"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected bool ShouldIgnore(ICellInfo c, ParsingContext context)
        {
            if(CellStateHelper.ShouldIgnore(IgnoreHiddenValues, c, context))
            {
                return true;
            }
            if(IgnoreErrors && c.IsExcelError)
            {
                return true;
            }
            return false;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arg"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected bool ShouldIgnore(FunctionArgument arg, ParsingContext context)
        {
            if (CellStateHelper.ShouldIgnore(IgnoreHiddenValues, arg, context))
            {
                return true;
            }
            if(IgnoreErrors && arg.ValueIsExcelError)
            {
                return true;
            }
            return false;
        }

    }
}
