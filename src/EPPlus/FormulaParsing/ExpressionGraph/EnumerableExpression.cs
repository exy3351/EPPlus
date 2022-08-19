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
    public class EnumerableExpression : Expression
    {
        /// <summary>
        /// 
        /// </summary>
        public EnumerableExpression()
            : this(new ExpressionCompiler())
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expressionCompiler"></param>
        public EnumerableExpression(IExpressionCompiler expressionCompiler)
        {
            _expressionCompiler = expressionCompiler;
        }

        private readonly IExpressionCompiler _expressionCompiler;
        /// <inheritdoc/>
        public override bool IsGroupedExpression
        {
            get { return false; }
        }
        /// <inheritdoc/>
        public override Expression PrepareForNextChild()
        {
            return this;
        }
        /// <inheritdoc/>
        public override CompileResult Compile()
        {
            var result = new List<object>();
            foreach (var childExpression in Children)
            {
                result.Add(_expressionCompiler.Compile(new List<Expression>{ childExpression }).Result);
            }
            return new CompileResult(result, DataType.Enumerable);
        }
    }
}
