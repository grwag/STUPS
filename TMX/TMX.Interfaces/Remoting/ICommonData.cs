﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 9/10/2014
 * Time: 9:32 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace Tmx.Interfaces.Remoting
{
    using System.Collections.Generic;
    
    /// <summary>
    /// Description of ICommonData.
    /// </summary>
    public interface ICommonData
    {
        IDictionary<string, object> Data { get; set; }
        void AddOrUpdateDataItem(ICommonDataItem commonDataItem);
    }
}
