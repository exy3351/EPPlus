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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// 
    /// </summary>
    public class FunctionArgument
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        public FunctionArgument(object val)
        {
            Value = val;
            DataType = DataType.Unknown;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        /// <param name="dataType"></param>
        public FunctionArgument(object val, DataType dataType)
            :this(val)
        {
            DataType = dataType;
        }

        private ExcelCellState _excelCellState;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="state"></param>
        public void SetExcelStateFlag(ExcelCellState state)
        {
            _excelCellState |= state;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        public bool ExcelStateFlagIsSet(ExcelCellState state)
        {
            return (_excelCellState & state) != 0;
        }
        /// <summary>
        /// 
        /// </summary>
        public object Value { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public DataType DataType { get; }
        /// <summary>
        /// 
        /// </summary>
        public Type Type
        {
            get { return Value != null ? Value.GetType() : null; }
        }
        /// <summary>
        /// 
        /// </summary>
        public int ExcelAddressReferenceId { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public bool IsExcelRange
        {
            get { return Value != null && Value is IRangeInfo; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool IsEnumerableOfFuncArgs
        {
            get { return Value != null && Value is IEnumerable<FunctionArgument>; }
        }
        /// <summary>
        /// 
        /// </summary>
        public IEnumerable<FunctionArgument> ValueAsEnumerableOfFuncArgs
        {
            get { return Value as IEnumerable<FunctionArgument>; }
        }
        /// <summary>
        /// 
        /// </summary>
        public bool ValueIsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(Value); }
        }
        /// <summary>
        /// 
        /// </summary>
        public ExcelErrorValue ValueAsExcelErrorValue
        {
            get { return ExcelErrorValue.Parse(Value.ToString()); }
        }
        /// <summary>
        /// 
        /// </summary>
        public IRangeInfo ValueAsRangeInfo
        {
            get { return Value as IRangeInfo; }
        }
        /// <summary>
        /// 
        /// </summary>
        public object ValueFirst
        {
            get
            {
                if (Value is INameInfo)
                {
                    Value = ((INameInfo)Value).Value;
                }
                var v = Value as IRangeInfo;
                if (v==null)
                {
                    return Value;
                }
                else
                {
                    return v.GetValue(v.Address._fromRow, v.Address._fromCol);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        public string ValueFirstString
        {
            get
            {
                var v = ValueFirst;
                if (v == null) return default(string);
                return ValueFirst.ToString();
            }
        }

    }
}
