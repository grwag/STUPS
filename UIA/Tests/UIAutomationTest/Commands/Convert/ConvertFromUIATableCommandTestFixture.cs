﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 08.12.2011
 * Time: 11:17
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace UIAutomationTest.Commands.Convert
{
    using System.Management.Automation;

    /// <summary>
    /// Description of ConvertFromUiaTableCommandTestFixture.
    /// </summary>
    [MbUnit.Framework.TestFixture][NUnit.Framework.TestFixture] // [TestFixture(Description="ConvertFrom-UiaTableCommand test")]
    public class ConvertFromUiaTableCommandTestFixture
    {
        [MbUnit.Framework.SetUp][NUnit.Framework.SetUp]
        public void PrepareRunspace()
        {
            MiddleLevelCode.PrepareRunspace();
        }
        
        [MbUnit.Framework.Test][NUnit.Framework.Test] //[Test(Description="InputObject ProcessRecord test Null via pipeline")]
        [MbUnit.Framework.Category("Slow")]
        [MbUnit.Framework.Category("NoForms")]
        public void ConvertFromUiaTable_TestPipelineInput()
        {
            string codeSnippet = 
                @"if ( ($null | ConvertFrom-UiaTable) ) { 1; }else{ 0; }";
            System.Collections.ObjectModel.Collection<PSObject> coll = 
                CmdletUnitTest.TestRunspace.RunPSCode(codeSnippet);
            MbUnit.Framework.Assert.IsTrue(coll[0].ToString() == "0");
        }
        
        [MbUnit.Framework.Test][NUnit.Framework.Test] //[Test(Description="ProcessRecord test Null via parameter")]
        [MbUnit.Framework.Category("Slow")]
        [MbUnit.Framework.Category("NoForms")]
        public void ConvertFromUiaTable_TestParameterInputNull()
        {
//            string codeSnippet = 
//                @"if ((ConvertFrom-UiaTable -InputObject $null)) { 1; }else{ 0; }";
//            System.Collections.ObjectModel.Collection<PSObject> coll =
//                CmdletUnitTest.TestRunspace.RunPSCode(codeSnippet);
//            Assert.IsNull(coll);
            
            CmdletUnitTest.TestRunspace.RunAndGetTheException(
                @"if ((ConvertFrom-UiaTable -InputObject $null)) { 1; } else { 0; }",
                "ParameterBindingValidationException",
                @"Cannot validate argument on parameter 'InputObject'. The argument is null or empty. Supply an argument that is not null or empty and then try the command again.");
            
//            UIAutomationTest.Commands.Convert.ConvertFromUiaTableCommandTestFixture.TestParameterInputNull:
//System.Management.Automation.ParameterBindingValidationException : Cannot validate argument on parameter 'InputObject'. The argument is null or empty. Supply an argument that is not null or empty and then try the command again.
//  ----> System.Management.Automation.ValidationMetadataException : The argument is null or empty. Supply an argument that is not null or empty and then try the command again.
        }
        
        [MbUnit.Framework.Test][NUnit.Framework.Test] //[Test(Description="ProcessRecord test Is Not AutomationElement")]
        [MbUnit.Framework.Category("Slow")]
        [MbUnit.Framework.Category("NoForms")]
        public void ConvertFromUiaTable_TestParameterInputOtherType()
        {
//            string codeSnippet = 
//                @"if ((ConvertFrom-UiaTable -InputObject (New-Object System.Windows.forms.Label))) { 1; }else{ 0; }";
//            System.Collections.ObjectModel.Collection<PSObject> coll =
//                CmdletUnitTest.TestRunspace.RunPSCode(codeSnippet);
//            Assert.IsNull(coll);
            
            CmdletUnitTest.TestRunspace.RunAndGetTheException(
                @"if ((ConvertFrom-UiaTable -InputObject (New-Object System.Windows.forms.Label))) { 1; } else { 0; }",
                "ParameterBindingException",
                //@"Cannot bind parameter 'InputObject'. Cannot convert the ""System.Windows.Forms.Label, Text: "" value of type ""System.Windows.Forms.Label"" to type ""System.Windows.Automation.AutomationElement"".");
                @"Cannot bind parameter 'InputObject'. Cannot convert the ""System.Windows.Forms.Label, Text: "" value of type ""System.Windows.Forms.Label"" to type ""UIAutomation.IUiElement"".");
            
//            UIAutomationTest.Commands.Convert.ConvertFromUiaTableCommandTestFixture.TestParameterInputOtherType:
//System.Management.Automation.ParameterBindingException : Cannot bind parameter 'InputObject'. Cannot convert the "System.Windows.Forms.Label, Text: " value of type "System.Windows.Forms.Label" to type "System.Windows.Automation.AutomationElement".
//  ----> System.Management.Automation.PSInvalidCastException : Cannot convert the "System.Windows.Forms.Label, Text: " value of type "System.Windows.Forms.Label" to type "System.Windows.Automation.AutomationElement".
        }
        
        [MbUnit.Framework.TearDown][NUnit.Framework.TearDown]
        public void DisposeRunspace()
        {
            MiddleLevelCode.DisposeRunspace();
        }
        
    }
    
    

}