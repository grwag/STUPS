﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 9/3/2012
 * Time: 11:55 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace Tmx
{
    using System.Management.Automation;
    
    /// <summary>
    /// Description of TestCaseCmdletBase.
    /// </summary>
    public class TestCaseCmdletBase : DatabaseCmdletBase
    {
        // 20141211
        public TestCaseCmdletBase()
        {
            TestCode = new ScriptBlock[]{};
        }
        
        #region Parameters
        [Parameter(Mandatory = false)]
        public ScriptBlock[] TestCode { get; set; }
        #endregion Parameters
    }
}
