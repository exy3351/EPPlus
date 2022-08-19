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
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// 
    /// </summary>
    public class DateExpression : AtomicExpression
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        public DateExpression(string expression)
            : base(expression)
        {

        }
        /// <inheritdoc/>
        public override CompileResult Compile()
        {
            var date = double.Parse(ExpressionString);
            return new CompileResult(DateTime.FromOADate(date), DataType.Date);
        }
    }
}
