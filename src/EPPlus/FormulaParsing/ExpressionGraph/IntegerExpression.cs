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
    public class IntegerExpression : AtomicExpression
    {
        private double? _compiledValue;
        private bool _negate;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        public IntegerExpression(string expression)
            : this(expression, false)
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="negate"></param>
        public IntegerExpression(string expression, bool negate)
            : base(expression)
        {
            _negate = negate;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        public IntegerExpression(double val)
            : base(val.ToString(CultureInfo.InvariantCulture))
        {
            _compiledValue = Math.Floor(val);
        }
        /// <inheritdoc/>
        public override CompileResult Compile()
        {
            double result = _compiledValue ?? double.Parse(ExpressionString, CultureInfo.InvariantCulture);
            result = _negate ? result * -1 : result;
            return new CompileResult(result, DataType.Integer);
        }
    }
}
