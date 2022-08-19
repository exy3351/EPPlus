/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           EPPlus 6
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Information and help methods about a cell
    /// </summary>
    public interface ICellInfo
    {
        /// <summary>
        /// 
        /// </summary>
        string Address { get; }

        /// <summary>
        /// 
        /// </summary>
        string WorksheetName { get; }
        /// <summary>
        /// 
        /// </summary>
        int Row { get; }
        /// <summary>
        /// 
        /// </summary>
        int Column { get; }

        /// <summary>
        /// 
        /// </summary>
        ulong Id { get; }
        /// <summary>
        /// 
        /// </summary>
        string Formula { get; }
        /// <summary>
        /// 
        /// </summary>
        object Value { get; }
        /// <summary>
        /// 
        /// </summary>
        double ValueDouble { get; }
        /// <summary>
        /// 
        /// </summary>
        double ValueDoubleLogical { get; }
        /// <summary>
        /// 
        /// </summary>
        bool IsHiddenRow { get; }
        /// <summary>
        /// 
        /// </summary>
        bool IsExcelError { get; }
        /// <summary>
        /// 
        /// </summary>
        IList<Token> Tokens { get; }
    }
}
