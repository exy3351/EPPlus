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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    /// <summary>
    /// 
    /// </summary>
    public class StringConcatStrategy : CompileStrategy
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        public StringConcatStrategy(Expression expression)
            : base(expression)
        {
           
        }
        /// <inheritdoc/>
        public override Expression Compile()
        {
            var newExp = _expression is ExcelAddressExpression ? _expression : ExpressionConverter.Instance.ToStringExpression(_expression);
            newExp.Prev = _expression.Prev;
            newExp.Next = _expression.Next;
            if (_expression.Prev != null)
            {
                _expression.Prev.Next = newExp;
            }
            if (_expression.Next != null)
            {
                _expression.Next.Prev = newExp;
            }
            return newExp.MergeWithNext();
        }
    }
}
