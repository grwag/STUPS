/*
 * Created by SharpDevelop.
 * User: Alexander Petrovskiy
 * Date: 11/22/2012
 * Time: 11:39 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

namespace SePSX
{
    using System;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Firefox;
    using OpenQA.Selenium.IE;
    using OpenQA.Selenium.Remote;

    using Commands;

    //using Ninject;


    using Autofac;
    using Autofac.Builder;

    /// <summary>
    /// Description of WebDriverFactory.
    /// </summary>
    //public static class WebDriverFactory
    public static class WebDriverFactory // : IWebDriverFactory
    {
        static WebDriverFactory()
        {
            try
            {
                AutofacModule = new WebDriverModule();
            }
            catch (Exception eLoadingModule)
            {
                Console.WriteLine("Loading of module failed; " + eLoadingModule.Message);
            }
        }

        private static Module _autofacModule;
        internal static Module AutofacModule
        {
            get { return _autofacModule; }
            set { _autofacModule = value; _initFlag = false; }
        }

        private static ContainerBuilder _builder;
        internal static IContainer Container;
        private static bool _initFlag = false;

        internal static void Init()
        {
            if (!_initFlag)
            {
                try
                {
                    _builder = new ContainerBuilder();

                    _builder.RegisterModule(AutofacModule);

                    Container = null;

                    var container = _builder.Build(ContainerBuildOptions.Default);

                    Container = container;

                }
                catch (Exception efgh)
                {
                    Console.WriteLine(efgh.Message);
                }

                _initFlag = true;
            }
        }

        private static IWebDriver _driver;

        public static IWebDriver GetDriver(StartDriverCmdletBase cmdlet, Drivers driverType)
        {

            try
            {

                // enumerate driver processes before creating new one
                SeHelper.CollectDriverProcesses(driverType);

                switch (driverType)
                {
                    case Drivers.Firefox:

                        var ffCapabilities = CapabilitiesFactory.GetCapabilities();

                        _driver = new FirefoxDriver(new FirefoxOptions());


                        SeHelper.GetDriverProcess(Drivers.Firefox, _driver.Title + SeHelper.DriverTitleComplementFirefox.Substring(3));

                        _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
                        break;
                    case Drivers.Safari:
                        _driver = GetNativeDriver(driverType);

                        SeHelper.GetDriverProcess(Drivers.Safari, _driver.Title + SeHelper.DriverTitleComplementSafari);
                        break;
                    case Drivers.Html:
                    default:
                        throw new Exception("Invalid value for Drivers");
                }

                return _driver;

            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                return null;
            }
        }

        private static IWebDriver _nativeDriver;

        internal static IWebDriver NativeDriver
        {
            get { return _nativeDriver; }
            set { _nativeDriver = value; }
        }

        internal static IWebDriver GetNativeDriver(Drivers driverType)
        {
            try
            {
                IWebDriver nativeDriver = null;

                switch (driverType)
                {
                    case Drivers.Chrome:

                        break;
                    case Drivers.Firefox:

                        break;
                    case Drivers.InternetExplorer:

                        break;
                    case Drivers.Safari:

                        nativeDriver = NativeDriver;

                        break;
                    case Drivers.Html:

                        break;
                    default:
                        throw new Exception("Invalid value for Drivers");
                }

                return nativeDriver;
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
                return null;
            }
        }

        public static IWebDriver GetDriverServer(StartSeWebDriverCommand cmdlet)
        {
            IWebDriver driver = null;
            var driverDirectoryPath = string.Empty;
            ChromeOptions chromeOptions = null;
            InternetExplorerOptions ieOptions = null;
            var commandTimeout = TimeSpan.FromSeconds(60.0);
            DriverService service = null;
            var listOfParameters =
                new System.Collections.Generic.List<Autofac.Core.Parameter>();

            // determine the type of dirver server
            var driverServerType = DriverServers.None;
            var ieArchitecture = InternetExplorer.X86;

            switch (cmdlet.DriverType)
            {
                case Drivers.Chrome:
                    cmdlet.WriteVerbose(cmdlet, "required ChromeDriver");
                    driverServerType = DriverServers.Chrome;
                    //driverType = Drivers.Chrome;
                    break;
                case Drivers.Firefox:

                    break;
                case Drivers.InternetExplorer:
                    cmdlet.WriteVerbose(cmdlet, "required InternetExplorerDriver");
                    driverServerType = DriverServers.Ie;
                    ieArchitecture = cmdlet.Architecture;
                    //driverType = Drivers.InternetExplorer;
                    break;
                case Drivers.Safari:

                    break;
                case Drivers.Html:

                    break;
                default:
                    throw new Exception("Invalid value for Drivers");
            }

            // collect processes before running the server
            cmdlet.WriteVerbose(cmdlet, "collect processes");

            SeHelper.CollectDriverProcesses(cmdlet.DriverType);

            if (string.IsNullOrEmpty(cmdlet.DriverDirectoryPath))
            {

                cmdlet.WriteVerbose(cmdlet, "using the default driver directory path");
                driverDirectoryPath =
                    System.IO.Path.GetDirectoryName(cmdlet.GetType().Assembly.Location);
                if (DriverServers.Ie == driverServerType && InternetExplorer.X86 == ieArchitecture)
                {
                    driverDirectoryPath += "\\32\\";
                }
                if (DriverServers.Ie == driverServerType && InternetExplorer.X64 == ieArchitecture)
                {
                    driverDirectoryPath += "\\64\\";
                }
            }
            else
            {

                cmdlet.WriteVerbose(cmdlet, "using the path from the cmdlet parameter");
                driverDirectoryPath = cmdlet.DriverDirectoryPath;
            }
            cmdlet.WriteVerbose(cmdlet, driverDirectoryPath);

            if (DriverServers.Chrome == driverServerType)
            {

                if (null == cmdlet.ChromeOptions)
                {

                    cmdlet.WriteVerbose(cmdlet, "using the default ChromeOptions");

                    chromeOptions = new ChromeOptions();
                }
                else
                {

                    cmdlet.WriteVerbose(cmdlet, "using the supplied ChromeOptions");
                    chromeOptions = cmdlet.ChromeOptions;
                }
            }
            if (DriverServers.Ie == driverServerType)
            {

                if (null == cmdlet.IeOptions)
                {

                    cmdlet.WriteVerbose(cmdlet, "using the default InternetExplorerOptions");

                    ieOptions = new InternetExplorerOptions();
                    ieOptions.IgnoreZoomLevel = true;
                }
                else
                {

                    cmdlet.WriteVerbose(cmdlet, "using the supplied InternetExplorerOptions");
                    ieOptions = cmdlet.IeOptions;
                }
            }

            if (null != cmdlet.Timeout && 0 != cmdlet.Timeout && Preferences.Timeout != cmdlet.Timeout)
            { // ??

                cmdlet.WriteVerbose(cmdlet, "setting the commandTimeout");
                commandTimeout = TimeSpan.FromMilliseconds(cmdlet.Timeout);
                cmdlet.WriteVerbose(cmdlet, "commandTimeout = " + commandTimeout.ToString());
            }

            cmdlet.WriteVerbose(cmdlet, "creating a DriverService");
            switch (driverServerType)
            {
                //                case DriverServers.none:
                //                    
                //                    break;
                case DriverServers.Chrome:
                    //Console.WriteLine("driver server 00012c");
                    cmdlet.WriteVerbose(cmdlet, "creating a ChromeDriverService");
                    //Console.WriteLine("creating chrome driver service");
                    service = ChromeDriverService.CreateDefaultService(driverDirectoryPath);
                    //Console.WriteLine("the chrome driver service has been created");
                    break;
                case DriverServers.Ie:

                    cmdlet.WriteVerbose(cmdlet, "creating a InternetExplorerDriverService");

                    service = InternetExplorerDriverService.CreateDefaultService(driverDirectoryPath);

                    break;
                default:
                    throw new Exception("Invalid value for DriverServers");
            }

            switch (driverServerType)
            {
                case DriverServers.Chrome:

                    listOfParameters.Add(new NamedParameter("service", service));
                    listOfParameters.Add(new NamedParameter("options", chromeOptions));
                    listOfParameters.Add(new NamedParameter("commandTimeout", commandTimeout));

                    cmdlet.WriteVerbose(cmdlet, "creating the ChromeDriver");
                    try
                    {
                        driver = new ChromeDriver((ChromeDriverService)service, chromeOptions, commandTimeout);
                    }
                    catch (Exception eChDrv)
                    {
                        Console.WriteLine(eChDrv.Message);
                        Console.WriteLine(eChDrv.StackTrace);
                    }

                    break;
                case DriverServers.Ie:
                    listOfParameters.Add(new NamedParameter("service", service));
                    listOfParameters.Add(new NamedParameter("options", ieOptions));
                    listOfParameters.Add(new NamedParameter("commandTimeout", commandTimeout));
                    cmdlet.WriteVerbose(cmdlet, "creating the InternetExplorerDriver");
                    driver =
                        Container.ResolveNamed<IWebDriver>(
                            driverServerType.ToString(),
                            listOfParameters.ToArray());
                    break;
                default:

                    throw new Exception("Invalid value for DriverServers");
            }

            cmdlet.WriteVerbose(cmdlet, "findng out the process of the driver created");
            switch (cmdlet.DriverType)
            {
                case Drivers.Chrome:
                    SeHelper.GetDriverProcess(cmdlet.DriverType, driver.Title + SeHelper.DriverTitleComplementChrome);
                    break;
                case Drivers.InternetExplorer:
                    SeHelper.GetDriverProcess(cmdlet.DriverType, driver.Title + SeHelper.DriverTitleComplementInternetExplorer);
                    break;
            }

            return driver;

        }

        public static FirefoxProfile GetFirefoxProfile(FirefoxProfileCmdletBase cmdlet)
        {
            FirefoxProfile profile = null;
            var listOfParameters =
                new System.Collections.Generic.List<Autofac.Core.Parameter>();

            var profileDirectory =
                ((NewSeFirefoxProfileCommand)cmdlet).ProfileDirectoryPath;
            bool deleteSourceOnClean =
                ((NewSeFirefoxProfileCommand)cmdlet).DeleteSource;

            if (!string.IsNullOrEmpty(profileDirectory))
            {
                listOfParameters.Add(
                    new NamedParameter(
                        "profileDirectory",
                        profileDirectory));

                if (deleteSourceOnClean)
                {
                    listOfParameters.Add(
                        new NamedParameter(
                            "deleteSourceOnClean",
                            deleteSourceOnClean));

                    profile =
                        Container.ResolveNamed<FirefoxProfile>(
                            FirefoxProfileConstructorOptions.FfWithPathAndBool.ToString(),
                            listOfParameters);
                }
                else
                {
                    profile =
                        Container.ResolveNamed<FirefoxProfile>(
                            FirefoxProfileConstructorOptions.FfWithPath.ToString(),
                            listOfParameters);
                }
            }
            else
            {
                profile =
                    Container.ResolveNamed<FirefoxProfile>(
                        FirefoxProfileConstructorOptions.FfBare.ToString());
            }

            return profile;
        }

        public static IWebDriver GetFirefoxDriver(StartSeDriverServerCommand cmdlet)
        {
            return (new FirefoxDriver());
        }
    }
}
