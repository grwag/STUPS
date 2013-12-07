﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 29.11.2011
 * Time: 14:17
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;

namespace UIAutomation
{
    //using System;
    using System.Management.Automation;
    //using System.Windows.Automation;
    
    using System.Linq;

    /// <summary>
    /// Description of WaitCmdletBase.
    /// </summary>
    public class WaitCmdletBase : HasTimeoutCmdletBase
    {
        #region Constructor
        public WaitCmdletBase()
        {
            // duplicated !!!
            InputObject = null;
        }
        #endregion Constructor

        #region Parameters
        #endregion Parameters
        
        protected void WaitIfCondition(
            // 20131109
            //AutomationElement _control,
            IUiElement _control,
            bool isEnabledOrIsVisible)
        {
            // 20131109
            //_control = this.InputObject[0];
            _control = InputObject.Cast<IUiElement>().ToArray()[0];
            
            if (isEnabledOrIsVisible) {
                Wait = !(_control.Current).IsEnabled;
            } else {
                Wait = (_control.Current).IsOffscreen;
            }
            do
            {
                SleepAndRunScriptBlocks(this);
                
                DateTime nowDate = DateTime.Now;
                try {
                    string tempIsReport = string.Empty;
                    tempIsReport = isEnabledOrIsVisible ? _control.Current.IsEnabled.ToString() : _control.Current.IsOffscreen.ToString();

                    /*
                    if (isEnabledOrIsVisible) {
                        tempIsReport = _control.Current.IsEnabled.ToString();
                    } else {
                        tempIsReport = _control.Current.IsOffscreen.ToString();
                    }
                    */
                    WriteVerbose(this,
                                 "AutomationID: " + 
                                 _control.Current.AutomationId +
                                 ", title: " + 
                                 _control.Current.Name + 
                                 ", Enabled = " + 
                                 tempIsReport +
                                 ", seconds: " + 
                                 ((nowDate - StartDate).TotalSeconds).ToString());
                } catch { }
                if (!CheckAndPrepareInput(this))
                {
                    WriteObject(this, false);
                    
                    WriteError(
                        this,
                        "An unknown error while checking the control.",
                        "CheckingControl",
                        ErrorCategory.InvalidResult,
                        true);
                    
                    return;
                }
                
                if (isEnabledOrIsVisible) {
                    Wait = !(_control.Current).IsEnabled;
                } else {
                    Wait = (_control.Current).IsOffscreen;
                }
                if ((nowDate - StartDate).TotalSeconds > Timeout / 1000)
                {
                    WriteVerbose(this, "timeout expired for AutomationId: " + 
                                 _control.Current.AutomationId +
                                ", title: " +
                                 _control.Current.Name);
                    
                    WriteError(
                        this,
                        CmdletName(this) + ": timeout expired for AutomationId: " + 
                        _control.Current.AutomationId +
                        ", title: " +
                        _control.Current.Name,
                        "TimeoutExpired",
                        ErrorCategory.OperationTimeout,
                        true);
                }
                if (_control != null) continue;
                
                /*
                // code is heuristically unavailable
                WriteVerbose(this, "the control is unavailable");
                    
                this.WriteError(
                    this,
                    CmdletName(this) + ": control is unavailable. AutomationId: " + 
                    _control.Current.AutomationId +
                    ", title: " +
                    _control.Current.Name,
                    "TimeoutExpired",
                    ErrorCategory.OperationTimeout,
                    true);
                */

                /*
                if (_control == null) {
                    WriteVerbose(this, "the control is unavailable");
                    
                    this.WriteError(
                        this,
                        CmdletName(this) + ": control is unavailable. AutomationId: " + 
                        _control.Current.AutomationId +
                        ", title: " +
                        _control.Current.Name,
                        "TimeoutExpired",
                        ErrorCategory.OperationTimeout,
                        true);
                }
                */
            } while (Wait);
        }
    }
}
