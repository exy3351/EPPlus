using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelRangeExpression : Expression
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rangeInfo"></param>
        public ExcelRangeExpression(IRangeInfo rangeInfo)
        {
            _rangeInfo = rangeInfo;
        }

        private readonly IRangeInfo _rangeInfo;
        /// <inheritdoc/>
        public override bool IsGroupedExpression => false;
        /// <inheritdoc/>
        public override CompileResult Compile()
        {
            return new CompileResult(_rangeInfo, DataType.Enumerable);
        }
    }
}
