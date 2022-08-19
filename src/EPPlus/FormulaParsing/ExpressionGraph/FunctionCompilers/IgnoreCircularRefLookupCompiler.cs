using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    /// <summary>
    /// 
    /// </summary>
    public class IgnoreCircularRefLookupCompiler : LookupFunctionCompiler
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="function"></param>
        /// <param name="context"></param>
        public IgnoreCircularRefLookupCompiler(ExcelFunction function, ParsingContext context) : base(function, context)
        {
        }
        /// <inheritdoc/>
        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            foreach(var child in children)
            {
                child.IgnoreCircularReference = true;
            }
            return base.Compile(children);
        }
    }
}
