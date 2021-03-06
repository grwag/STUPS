﻿/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 12/3/2012
 * Time: 3:36 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace SePSXUnitTests.Commands.Navigation
{
    using SePSX;
    using SePSX.Commands;
    using MbUnit.Framework;
    using Autofac;
    
    /// <summary>
    /// Description of InvokeSeNavigateForwardCommandTestFixture.
    /// </summary>
    [TestFixture]
    public class InvokeSeNavigateForwardCommandTestFixture
    {
        public InvokeSeNavigateForwardCommandTestFixture()
        {
        }
        
        [SetUp]
        public void SetUp()
        {
            UnitTestingHelper.PrepareUnitTestDataStore();
        }
        
        [TearDown]
        public void TearDown()
        {
        }
        
        private void enterUrl(string firstUrl, string secondUrl)
        {
            StartSeChromeCommand cmdlet0 =
                WebDriverFactory.Container.Resolve<StartSeChromeCommand>();
            SeStartChromeCommand command0 =
                new SeStartChromeCommand(cmdlet0);
            command0.Execute();
            
            EnterSeUrlCommand cmdlet1 =
                WebDriverFactory.Container.Resolve<EnterSeUrlCommand>();
            cmdlet1.InputObject =
                new FakeWebDriver[]{ ((FakeWebDriver)(object)PSTestLib.UnitTestOutput.LastOutput[0]) };
            cmdlet1.Url = firstUrl;
            SeEnterUrlCommand command1 =
                new SeEnterUrlCommand(cmdlet1);
            command1.Execute();

            EnterSeUrlCommand cmdlet2 =
                WebDriverFactory.Container.Resolve<EnterSeUrlCommand>();
            cmdlet2.InputObject =
                new FakeWebDriver[]{ ((FakeWebDriver)(object)PSTestLib.UnitTestOutput.LastOutput[0]) };
            cmdlet2.Url = secondUrl;
            SeEnterUrlCommand command2 =
                new SeEnterUrlCommand(cmdlet2);
            command2.Execute();

            InvokeSeNavigateBackCommand cmdlet3 =
                WebDriverFactory.Container.Resolve<InvokeSeNavigateBackCommand>();
            cmdlet3.InputObject =
                new FakeWebDriver[]{ ((FakeWebDriver)(object)PSTestLib.UnitTestOutput.LastOutput[0]) };
            SeInvokeNavigateBackCommand command3 =
                new SeInvokeNavigateBackCommand(cmdlet3);
            command3.Execute();

            InvokeSeNavigateForwardCommand cmdlet4 =
                WebDriverFactory.Container.Resolve<InvokeSeNavigateForwardCommand>();
            cmdlet4.InputObject =
                new FakeWebDriver[]{ ((FakeWebDriver)(object)PSTestLib.UnitTestOutput.LastOutput[0]) };
            SeInvokeNavigateForwardCommand command4 =
                new SeInvokeNavigateForwardCommand(cmdlet4);
            command4.Execute();
        }
        
        [Test]
        [Category("Fast")]
        public void Url_Right_Web()
        {
            string secondUrl = @"http://google.com";
            string firstUrl = @"http://yahoo.com";
            enterUrl(firstUrl, secondUrl);
            Assert.AreEqual(
                secondUrl,
                ((string)(object)PSTestLib.UnitTestOutput.LastOutput[0]));
        }
        
        [Test]
        [Category("Fast")]
        public void Url_Right_File()
        {
            string secondUrl = @"C:\1\1.txt";
            string firstUrl = @"C:\2\2.doc";
            enterUrl(firstUrl, secondUrl);
            Assert.AreEqual(
                secondUrl,
                ((string)(object)PSTestLib.UnitTestOutput.LastOutput[0]));
        }
        
        [Test]
        [Category("Fast")]
        public void Url_Wrong()
        {
            string secondUrl = "wrong url";
            string firstUrl = @"broken url";
            enterUrl(firstUrl, secondUrl);
            Assert.AreEqual(
                secondUrl,
                ((string)(object)PSTestLib.UnitTestOutput.LastOutput[0]));
        }
        
        [Test]
        [Category("Fast")]
        public void Url_Empty()
        {
            string secondUrl = string.Empty;
            string firstUrl = string.Empty;
            enterUrl(firstUrl, secondUrl);
            Assert.AreEqual(
                secondUrl,
                ((string)(object)PSTestLib.UnitTestOutput.LastOutput[0]));
        }
    }
}
