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
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// 
    /// </summary>
    public class ObjectEnumerableArgConverter : CollectionFlattener<object>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ignoreHidden"></param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public virtual IEnumerable<object> ConvertArgs(bool ignoreHidden, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
                {
                    if (arg.Value is IRangeInfo)
                    {
                        foreach (var cell in (IRangeInfo)arg.Value)
                        {
                            if (!CellStateHelper.ShouldIgnore(ignoreHidden, cell, context))
                            {
                                argList.Add(cell.Value);
                            }
                        }
                    }
                    else
                    {
                       argList.Add(arg.Value);
                    }
                })
            ;

        }
    }
}
