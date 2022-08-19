﻿/*************************************************************************************************
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

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// 
    /// </summary>
    public class MeasurementFont
    {
        /// <summary>
        /// 
        /// </summary>
        public string FontFamily { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public MeasurementFontStyles Style { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public float Size { get; set; }
    }
}
