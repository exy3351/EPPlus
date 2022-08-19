/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           EPPlus Interfaces 1.0
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// 
    /// </summary>
    [Flags]
    public enum MeasurementFontStyles
    {
        /// <summary>
        /// Normal text.
        /// </summary>
        
        Regular = 0,
        /// <summary>
        /// Bold text.
        /// </summary>
        Bold = 1,
        /// <summary>
        /// Italic text.
        /// </summary>
        Italic = 2,
        /// <summary>
        /// Underlined text.
        /// </summary>
        Underline = 4,
        /// <summary>
        ///  Text with a line through the middle.
        /// </summary>
        Strikeout = 8
    }
}
