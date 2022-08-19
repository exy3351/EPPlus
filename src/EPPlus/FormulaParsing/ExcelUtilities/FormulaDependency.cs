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
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// 
    /// </summary>
    public class FormulaDependency
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="scope"></param>
        public FormulaDependency(ParsingScope scope)
	    {   
            ScopeId = scope.ScopeId;
            Address = scope.Address;
	    }
        /// <summary>
        /// 
        /// </summary>
        public Guid ScopeId { get; private set; }
        /// <summary>
        /// 
        /// </summary>
        public RangeAddress Address { get; private set; }

        private List<RangeAddress> _referencedBy = new List<RangeAddress>();

        private List<RangeAddress> _references = new List<RangeAddress>();
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rangeAddress"></param>
        /// <exception cref="CircularReferenceException"></exception>
        public virtual void AddReferenceFrom(RangeAddress rangeAddress)
        {
            if (Address.CollidesWith(rangeAddress) || _references.Exists(x => x.CollidesWith(rangeAddress)))
            {
                throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
            }
            _referencedBy.Add(rangeAddress);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rangeAddress"></param>
        /// <exception cref="CircularReferenceException"></exception>
        public virtual void AddReferenceTo(RangeAddress rangeAddress)
        {
            if (Address.CollidesWith(rangeAddress) || _referencedBy.Exists(x => x.CollidesWith(rangeAddress)))
            {
                throw new CircularReferenceException("Circular reference detected at " + rangeAddress.ToString());
            }
            _references.Add(rangeAddress);
        }
    }
}
