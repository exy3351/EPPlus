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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// 
    /// </summary>
    public class BooleanExpression : AtomicExpression
    {
        private bool? _precompiledValue;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        public BooleanExpression(string expression)
            : base(expression)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        public BooleanExpression(bool value)
            : base(value ? "true" : "false")
        {
            _precompiledValue = value;
        }
        /// <inheritdoc/>
        public override CompileResult Compile()
        {
            var result = _precompiledValue ?? bool.Parse(ExpressionString);
            return new CompileResult(result, DataType.Boolean);
        }
    }
}
