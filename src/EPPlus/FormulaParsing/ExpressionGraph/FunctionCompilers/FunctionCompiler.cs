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
using System.Collections;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    /// <summary>
    /// 
    /// </summary>
    public abstract class FunctionCompiler
    {
        /// <summary>
        /// 
        /// </summary>
        protected ExcelFunction Function
        {
            get;
            private set;
        }
        /// <summary>
        /// 
        /// </summary>
        protected ParsingContext Context
        {
            get;
            private set;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="function"></param>
        /// <param name="context"></param>
        public FunctionCompiler(ExcelFunction function, ParsingContext context)
        {
            Require.That(function).Named("function").IsNotNull();
            Require.That(context).Named("context").IsNotNull();
            Function = function;
            Context = context;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="compileResult"></param>
        /// <param name="dataType"></param>
        /// <param name="args"></param>
        protected void BuildFunctionArguments(CompileResult compileResult, DataType dataType, List<FunctionArgument> args)
        {
            if (compileResult.Result is IEnumerable<object> && !(compileResult.Result is IRangeInfo))
            {
                var compileResultFactory = new CompileResultFactory();
                var argList = new List<FunctionArgument>();
                var objects = compileResult.Result as IEnumerable<object>;
                foreach (var arg in objects)
                {
                    var cr = compileResultFactory.Create(arg);
                    BuildFunctionArguments(cr, dataType, argList);
                }
                args.Add(new FunctionArgument(argList));
            }
            else
            {
                var funcArg = new FunctionArgument(compileResult.Result, dataType);
                funcArg.ExcelAddressReferenceId = compileResult.ExcelAddressReferenceId;
                if(compileResult.IsHiddenCell)
                {
                    funcArg.SetExcelStateFlag(Excel.ExcelCellState.HiddenCell);
                }
                args.Add(funcArg);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="args"></param>
        protected void BuildFunctionArguments(CompileResult result, List<FunctionArgument> args)
        {
            BuildFunctionArguments(result, result.DataType, args);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="children"></param>
        /// <returns></returns>
        public abstract CompileResult Compile(IEnumerable<Expression> children);
    }
}
