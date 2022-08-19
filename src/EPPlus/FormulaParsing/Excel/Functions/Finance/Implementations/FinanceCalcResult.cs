/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class FinanceCalcResult<T>
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        public FinanceCalcResult(T result)
        {
            Result = result;
            if(result is double)
            {
                DataType = DataType.Decimal;
            }
            else if(result is int)
            {
                DataType = DataType.Integer;
            }
            else if(result is System.DateTime)
            {
                DataType = DataType.Date;
            }
            else
            {
                DataType = DataType.Unknown;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        public FinanceCalcResult(T result, DataType dataType)
        {
            Result = result;
            DataType = dataType;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="error"></param>
        public FinanceCalcResult(eErrorType error)
        {
            HasError = true;
            ExcelErrorType = error;
        }
        /// <summary>
        /// 
        /// </summary>
        public T Result { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public DataType DataType { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public bool HasError
        {
            get; private set;
        }
        /// <summary>
        /// 
        /// </summary>
        public eErrorType ExcelErrorType { get; private set; }
    }
}
