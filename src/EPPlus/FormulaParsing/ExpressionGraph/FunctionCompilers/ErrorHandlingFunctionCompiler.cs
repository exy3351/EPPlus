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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    /// <summary>
    /// 
    /// </summary>
    public class ErrorHandlingFunctionCompiler : FunctionCompiler
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="function"></param>
        /// <param name="context"></param>
        public ErrorHandlingFunctionCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {

        }
        /// <inheritdoc/>
        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            foreach (var child in children)
            {
                try
                {
                    var arg = child.Compile();
                    BuildFunctionArguments(arg != null ? arg : null, args);
                }
                catch (ExcelErrorValueException efe)
                {
                    return ((ErrorHandlingFunction)Function).HandleError(efe.ErrorValue.ToString());
                }
                catch// (Exception e)
                {
                    return ((ErrorHandlingFunction)Function).HandleError(ExcelErrorValue.Values.Value);
                }
                
            }
            return Function.Execute(args, Context);
        }
    }
}
