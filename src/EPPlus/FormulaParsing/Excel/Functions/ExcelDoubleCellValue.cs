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
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// 
    /// </summary>
    [DebuggerDisplay("{Value}")]
    public struct ExcelDoubleCellValue : IComparable<ExcelDoubleCellValue>, IComparable
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        public ExcelDoubleCellValue(double val)
        {
            Value = val;
            CellRow = default(int?);
            CellCol = default(int?);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        /// <param name="cellRow"></param>
        /// <param name="cellCol"></param>
        public ExcelDoubleCellValue(double val, int cellRow, int cellCol)
        {
            Value = val;
            CellRow = cellRow;
            CellCol = cellCol;
        }
        /// <summary>
        /// 
        /// </summary>
        public int? CellRow;
        /// <summary>
        /// 
        /// </summary>
        public int? CellCol;
        /// <summary>
        /// 
        /// </summary>
        public double Value;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="d"></param>
        public static implicit operator double(ExcelDoubleCellValue d)
        {
            return d.Value;
        }
        /// <summary>
        ///  User-defined conversion from double to Digit
        /// </summary>
        /// <param name="d"></param>
        public static implicit operator ExcelDoubleCellValue(double d)
        {
            return new ExcelDoubleCellValue(d);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public int CompareTo(ExcelDoubleCellValue other)
        {
            return Value.CompareTo(other.Value);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public int CompareTo(object obj)
        {
            if(obj is double)
            {
                return Value.CompareTo((double)obj);
            }
            return Value.CompareTo(((ExcelDoubleCellValue)obj).Value);
        }
        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            return CompareTo(obj) == 0;
        }
        /// <inheritdoc/>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static bool operator ==(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
        {
            return a.Value.CompareTo(b.Value) == 0d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static bool operator ==(ExcelDoubleCellValue a, double b)
        {
            return a.Value.CompareTo(b) == 0d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static bool operator !=(ExcelDoubleCellValue a, ExcelDoubleCellValue b)
        {
            return a.Value.CompareTo(b.Value) != 0d;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static bool operator !=(ExcelDoubleCellValue a, double b)
        {
            return a.Value.CompareTo(b) != 0d;
        }
    }
}
