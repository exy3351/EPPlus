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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// 
    /// </summary>
    public interface INameInfo
    {
        /// <summary>
        /// 
        /// </summary>
        ulong Id { get; set; }
        /// <summary>
        /// 
        /// </summary>
        string Worksheet { get; set; }
        /// <summary>
        /// 
        /// </summary>
        string Name { get; set; }
        /// <summary>
        /// 
        /// </summary>
        string Formula { get; set; }
        /// <summary>
        /// 
        /// </summary>
        IList<Token> Tokens { get; }
        /// <summary>
        /// 
        /// </summary>
        object Value { get; set; }
    }
}
