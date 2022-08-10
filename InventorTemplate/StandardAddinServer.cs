using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.Helper;
using InventorTemplate.UI;
using Microsoft.Extensions.Configuration;
using NLog;
using NLog.Config;

namespace InventorTemplate
{
    [ProgId("InventorTemplate.StandardAddInServer")]
    [Guid(Globals.AddInClientId)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private UserInterfaceEvents _uiEvents;
        List<Ribbon> _ribbons = new List<Ribbon>();
        List<RibbonPanel> _ribbonPanels = new List<RibbonPanel>();
        List<RibbonTab> _ribbonTabs = new List<RibbonTab>();
        List<CommandControl> _buttons = new List<CommandControl>();
        List<ButtonDefinition> _buttonDefinitions = new List<ButtonDefinition>();

        private static ApplicationEvents _invAppEvents;
        public static ApplicationEvents InvAppEvents
        {
            get => _invAppEvents;
            set => _invAppEvents = value;
        }

        ButtonDefinition _defaultButton;
        ButtonDefinition _info;

        public UserInterfaceEvents UiEvents
	    {
		    [MethodImpl(MethodImplOptions.Synchronized)]
		    get => _uiEvents;

		    [MethodImpl(MethodImplOptions.Synchronized)]
		    set
		    {
			    if (_uiEvents != null)
			    {
				    _uiEvents.OnResetRibbonInterface -= UiEventsOnResetRibbonInterface;
			    }

			    _uiEvents = value;
			    if (_uiEvents != null)
			    {
				    _uiEvents.OnResetRibbonInterface += UiEventsOnResetRibbonInterface;
			    }
		    }
	    }

        /// <summary>
        /// This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for the first time. However, with the introduction of the ribbon this argument is always true.
        /// </summary>
        /// <param name="addInSiteObject">The add in site object.</param>
        /// <param name="firstTime">if set to <c>true</c> [first time].</param>
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false);

            IConfiguration config = builder.Build();

            LogManager.ThrowConfigExceptions = true;
            var logSettings = config.GetSection("Logging").Get<AppsettingsBinder>();
            var basePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            if (basePath != null)
                LogManager.Configuration = new XmlLoggingConfiguration(System.IO.Path.Combine(basePath, "nlog.config"));
            else
            {
                throw new ArgumentException(
                    $"nlog.config not found! Make sure there is a nlog config definition.");
            }

            LogManager.Configuration.Variables["logPath"] = logSettings.LogPath;
            var logger = LogManager.GetCurrentClassLogger();

            try
            {
                logger.Debug("Addin InventorTemplate Activated");


                Globals.InvApp = addInSiteObject.Application;
                Globals.InvApplicationAddInSite = addInSiteObject;
                UiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;
                InvAppEvents = Globals.InvApp.ApplicationEvents;
                InvAppEvents.OnApplicationOptionChange += InvAppEvents_OnApplicationOptionChange;

                var themeManager = Globals.InvApp.ThemeManager;
                var activeTheme = themeManager.ActiveTheme;
                string theme = activeTheme.Name;
                logger.Debug("Inventor ThemeManager ActiveTheme: " + theme);

                _info = UiDefinitionHelper.CreateButton("Info", "info", @"UI\ButtonResources\Info", theme);
                _defaultButton = UiDefinitionHelper.CreateButton("DefaultButton", "defaultButton", @"UI\ButtonResources\DefaultButton", theme);
                _buttonDefinitions.Add(_info);
                _buttonDefinitions.Add(_defaultButton);

                if (firstTime)
                    AddToUserInterface();
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Unexpected failure in the activation of the add-in ""InventorTemplate""" + System.Environment.NewLine + System.Environment.NewLine + ex.Message);
            }
        }

        public void Deactivate()
        {

            try
            {
                _buttonDefinitions = new List<ButtonDefinition>();
                _buttons = new List<CommandControl>();
                _ribbons = new List<Ribbon>();
                _uiEvents = null;

                foreach (var commandControl in _buttons)
                {
                    commandControl.Delete();
                }
            }
            catch
            {
                // ignored
            }

            if (_ribbonPanels != null)
            {
                for (int i = 0; i < _ribbonPanels.Count; i++)
                {
                    _ribbonPanels[i].Delete();
                    Marshal.ReleaseComObject(_ribbonPanels[i]);
                    _ribbonPanels[i] = null;
                }

            }

            _ribbonPanels = new List<RibbonPanel>();

            if (_ribbonTabs != null)
            {
                for (int i = 0; i < _ribbonTabs.Count; i++)
                {
                    _ribbonTabs[i].Delete();
                    Marshal.ReleaseComObject(_ribbonTabs[i]);
                    _ribbonTabs[i] = null;
                }

            }

            _ribbonTabs = new List<RibbonTab>();

            try
            {
                InvAppEvents.OnApplicationOptionChange -= InvAppEvents_OnApplicationOptionChange;
            }
            catch
            {
                // ignored
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public object Automation => null;
        public void ExecuteCommand(int commandId) { }

        private void AddToUserInterface()
        {
            var idwRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            var iptRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            var iamRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            var ipnRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Presentation"];
            _ribbons.Add(idwRibbon);
            _ribbons.Add(iptRibbon);
            _ribbons.Add(iamRibbon);
            _ribbons.Add(ipnRibbon);

            var tabIdw = UiDefinitionHelper.SetupTab("Testaddin", "Testaddin", idwRibbon);
            var tabIpt = UiDefinitionHelper.SetupTab("Testaddin", "Testaddin", iptRibbon);
            var tabIam = UiDefinitionHelper.SetupTab("Testaddin", "Testaddin", iamRibbon);
            var tabIpn = UiDefinitionHelper.SetupTab("Testaddin", "Testaddin", ipnRibbon);
            _ribbonTabs.Add(tabIdw);
            _ribbonTabs.Add(tabIpt);
            _ribbonTabs.Add(tabIam);
            _ribbonTabs.Add(tabIpn);

            var infoIdw = UiDefinitionHelper.SetupPanel("Info", "Info", tabIdw);
            var infoIpt = UiDefinitionHelper.SetupPanel("Info", "Info", tabIpt);
            var infoIam = UiDefinitionHelper.SetupPanel("Info", "Info", tabIam);
            var infoIpn = UiDefinitionHelper.SetupPanel("Info", "Info", tabIpn);
            _ribbonPanels.Add(infoIdw);
            _ribbonPanels.Add(infoIpt);
            _ribbonPanels.Add(infoIam);
            _ribbonPanels.Add(infoIpn);

            var addinPanelIdw = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIdw);
            var addinPanelIpt = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpt);
            var addinPanelIam = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIam);
            var addinPanelIpn = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpn);
            _ribbonPanels.Add(addinPanelIdw);
            _ribbonPanels.Add(addinPanelIpt);
            _ribbonPanels.Add(addinPanelIam);
            _ribbonPanels.Add(addinPanelIpn);

            if (_defaultButton != null)
            {
                var defaultButtonIdw = addinPanelIdw.CommandControls.AddButton(_defaultButton, true);
                var defaultButtonIpt = addinPanelIpt.CommandControls.AddButton(_defaultButton, true);
                var defaultButtonIam = addinPanelIam.CommandControls.AddButton(_defaultButton, true);
                var defaultButtonIpn = addinPanelIpn.CommandControls.AddButton(_defaultButton, true);
                _buttons.Add(defaultButtonIdw);
                _buttons.Add(defaultButtonIpt);
                _buttons.Add(defaultButtonIam);
                _buttons.Add(defaultButtonIpn);
            }
            if (_info != null)
            {
                var infoButtonIdw = infoIdw.CommandControls.AddButton(_info, true);
                var infoButtonIpt = infoIpt.CommandControls.AddButton(_info, true);
                var infoButtonIam = infoIam.CommandControls.AddButton(_info, true);
                var infoButtonIpn = infoIpn.CommandControls.AddButton(_info, true);
                _buttons.Add(infoButtonIdw);
                _buttons.Add(infoButtonIpt);
                _buttons.Add(infoButtonIam);
                _buttons.Add(infoButtonIpn);
            }
        }

        private void UiEventsOnResetRibbonInterface(NameValueMap context)
        {
	        AddToUserInterface();
        }

        private void InvAppEvents_OnApplicationOptionChange(EventTimingEnum BeforeOrAfter, NameValueMap Context, out HandlingCodeEnum HandlingCode)
        {
            if (BeforeOrAfter == EventTimingEnum.kAfter)
            {
                var themeManager = Globals.InvApp.ThemeManager;
                var activeTheme = themeManager.ActiveTheme;
                string theme = activeTheme.Name;

                // TODO create check if theme even changed
                // maybe create global theme variable

                //UiDefinitionHelper.SetButtonTheme(_buttonDefinitions, theme);
                Deactivate();
                Activate(Globals.InvApplicationAddInSite, true);
            }

            HandlingCode = HandlingCodeEnum.kEventNotHandled;
        }
    }
}
