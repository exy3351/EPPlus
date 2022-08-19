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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// 
    /// </summary>
    public abstract class CompileResultValidator
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="obj"></param>
        public abstract void Validate(object obj);

        private static CompileResultValidator _empty;
        /// <summary>
        /// 
        /// </summary>
        public static CompileResultValidator Empty
        {
            get { return _empty ?? (_empty = new EmptyCompileResultValidator()); }
        }
    }

    /// <summary>
    /// 
    /// </summary>
    internal class EmptyCompileResultValidator : CompileResultValidator
    {
        /// <inheritdoc/>
        public override void Validate(object obj)
        {
            // empty validator - do nothing
        }
    }
}
