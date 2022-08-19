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
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{

    /// <summary>
    /// 
    /// </summary>
    public abstract class Expression
    {
        internal string ExpressionString { get; private set; }
        private readonly List<Expression> _children = new List<Expression>();
        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<Expression> Children { get { return _children; } }
        /// <summary>
        /// 
        /// </summary>
        public Expression Next { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public Expression Prev { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public IOperator Operator { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public abstract bool IsGroupedExpression { get; }
        /// <summary>
        /// If set to true, <see cref="ExcelAddressExpression"></see>s that has a circular reference to their cell will be ignored when compiled
        /// </summary>
        public virtual bool IgnoreCircularReference
        {
            get; set;
        }

        /// <summary>
        /// 
        /// </summary>
        public Expression()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="expression"></param>
        public Expression(string expression)
        {
            ExpressionString = expression;
            Operator = null;
        }
        /// <summary>
        /// 
        /// </summary>
        public virtual bool HasChildren
        {
            get { return _children.Any(); }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public virtual Expression  PrepareForNextChild()
        {
            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="child"></param>
        /// <returns></returns>
        public virtual Expression AddChild(Expression child)
        {
            if (_children.Any())
            {
                var last = _children.Last();
                child.Prev = last;
                last.Next = child;
            }
            _children.Add(child);
            return child;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public virtual Expression MergeWithNext()
        {
            var expression = this;
            if (Next != null && Operator != null)
            {
                var result = Operator.Apply(Compile(), Next.Compile());
                expression = ExpressionConverter.Instance.FromCompileResult(result);
                if (expression is ExcelErrorExpression)
                {
                    expression.Next = null;
                    expression.Prev = null;
                    return expression;
                }
                if (Next != null)
                {
                    expression.Operator = Next.Operator;
                }
                else
                {
                    expression.Operator = null;
                }
                expression.Next = Next.Next;
                if (expression.Next != null) expression.Next.Prev = expression;
                expression.Prev = Prev;
            }
            else
            {
                throw (new FormatException("Invalid formula syntax. Operator missing expression."));
            }
            if (Prev != null)
            {
                Prev.Next = expression;
            }            
            return expression;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public abstract CompileResult Compile();

    }
}
