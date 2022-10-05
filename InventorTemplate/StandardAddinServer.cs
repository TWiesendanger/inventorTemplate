using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.UI;
using NLog;
using NLog.Config;
using Path = System.IO.Path;

namespace InventorTemplate
{
    [ProgId("InventorTemplate.StandardAddInServer")]
    [Guid(Globals.AddInClientId)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private UserInterfaceEvents _uiEvents;
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
            LogManager.ThrowConfigExceptions = true;
            var basePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (basePath != null)
                LogManager.Configuration = new XmlLoggingConfiguration(Path.Combine(basePath, "nlog.config"));
            else
            {
                throw new ArgumentException(
                    $"nlog.config not found! Make sure there is a nlog config definition.");
            }
            var logPath = Properties.Settings.Default.logPath;
            LogManager.Configuration.Variables["logPath"] = logPath;
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
                Globals.ActiveTheme = themeManager.ActiveTheme;
                string theme = Globals.ActiveTheme.Name;
                logger.Debug("Inventor ThemeManager ActiveTheme: " + theme);

                _info = UiDefinitionHelper.CreateButton("Info", "InventorTemplateInfo", @"UI\ButtonResources\Info", theme);
                _defaultButton = UiDefinitionHelper.CreateButton("DefaultButton", "InventorTemplateDefaultButton", @"UI\ButtonResources\DefaultButton", theme);
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
            ReleaseButtons();
            ReleaseRibbonPanels();
            ReleaseRibbonTabs();
            ReleaseAppEvents();

            try
            {
                _uiEvents = null;
                Marshal.ReleaseComObject(Globals.InvApp);
                Globals.InvApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
                // ignored
            }
        }
        private void ReleaseAppEvents()
        {
            try
            {
                InvAppEvents.OnApplicationOptionChange -= InvAppEvents_OnApplicationOptionChange;
            }
            catch
            {
                // ignored
            }
        }
        private void ReleaseRibbonTabs()
        {
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
        }
        private void ReleaseRibbonPanels()
        {
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
        }
        private void ReleaseButtons()
        {
            try
            {
                foreach (var buttonDefinition in _buttonDefinitions)
                {
                    buttonDefinition.Delete();
                    Marshal.ReleaseComObject(buttonDefinition);
                }

                foreach (var commandControl in _buttons)
                {
                    commandControl.Delete();
                    Marshal.ReleaseComObject(commandControl);
                }
            }
            catch
            {
                // ignored
            }

            _buttonDefinitions = new List<ButtonDefinition>();
            _buttons = new List<CommandControl>();
        }

        public object Automation => null;
        public void ExecuteCommand(int commandId) { }

        private void AddToUserInterface()
        {
            var idwRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            var iptRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            var iamRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            var ipnRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Presentation"];

            var tabIdw = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", idwRibbon);
            var tabIpt = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", iptRibbon);
            var tabIam = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", iamRibbon);
            var tabIpn = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", ipnRibbon);
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

        private void InvAppEvents_OnApplicationOptionChange(EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            if (beforeOrAfter == EventTimingEnum.kAfter)
            {
                var themeManager = Globals.InvApp.ThemeManager;
                var activeTheme = themeManager.ActiveTheme;
                string theme = activeTheme.Name;

                if (Globals.ActiveTheme.Name != theme) //check if theme has changed
                {
                    Deactivate();
                    Activate(Globals.InvApplicationAddInSite, true);
                }
            }

            handlingCode = HandlingCodeEnum.kEventNotHandled;
        }
    }
}