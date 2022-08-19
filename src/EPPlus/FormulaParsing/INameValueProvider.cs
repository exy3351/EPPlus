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

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// 
    /// </summary>
    public interface INameValueProvider
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        bool IsNamedValue(string key, string worksheet);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        object GetNamedValue(string key);
        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        object GetNamedValue(string key, string worksheet);
        /// <summary>
        /// 
        /// </summary>
        void Reload();
    }
}
