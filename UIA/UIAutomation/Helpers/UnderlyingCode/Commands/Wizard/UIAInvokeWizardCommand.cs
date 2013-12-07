﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 3/17/2013
 * Time: 2:44 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace UIAutomation
{
    using System;
    using System.Management.Automation;
    using Commands;
    
    /// <summary>
    /// Description of UiaInvokeWizardCommand.
    /// </summary>
    internal class UiaInvokeWizardCommand : UiaCommand
    {
        internal UiaInvokeWizardCommand(CommonCmdletBase cmdlet) : base (cmdlet)
        {
        }
        
        internal override void Execute()
        {
            WizardRunCmdletBase cmdlet =
                (WizardRunCmdletBase)Cmdlet;
            
            WizardHelper.InvokeWizard(cmdlet);
        }
    }
}
