using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.Helper;
using InventorTemplate.Helper.Logging;
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
            var logSettings = config.GetSection("Logging").Get<Logging>();
            LogManager.Configuration = new XmlLoggingConfiguration(@"C:\temp\sampleAddin\Testaddin\nlog.config");

            LogManager.Configuration.Variables["logPath"] = logSettings.LogPath;
            var logger = LogManager.GetCurrentClassLogger();

            try
            {
                logger.Debug("Addin InventorTemplate Activated");

                UiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;

                _info = UiDefinitionHelper.CreateButton("Info", "InventorTemplateInfo", @"UI\ButtonResources\InfoIcon");
                _defaultButton = UiDefinitionHelper.CreateButton("DefaultButton", "InventorTemplateDefaultButton", @"UI\ButtonResources\DefaultButton");

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
            _defaultButton = null;
            _info = null;
            //Globals.InvApp = null;
            _uiEvents = null;

            foreach (var ribbon in _ribbons)
            {
                foreach (RibbonTab ribbonTab in ribbon.RibbonTabs)
                {
                    if (ribbonTab.InternalName == "Testaddin")
                    {
                        foreach (RibbonPanel ribbonPanel in ribbonTab.RibbonPanels)
                        {
                            if (ribbonPanel.InternalName == "Info" || ribbonPanel.InternalName == "AddinCommands")
                            {
                                foreach (CommandControl commandControl in ribbonPanel.CommandControls)
                                {
                                    if (commandControl.InternalName.Contains("TestAddin"))
                                        commandControl.Delete();
                                }
                            }
                        }
                    }
                }
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

            _ribbonPanels = null;

            if (_ribbonTabs != null)
            {
                for (int i = 0; i < _ribbonTabs.Count; i++)
                {
                    _ribbonTabs[i].Delete();
                    Marshal.ReleaseComObject(_ribbonTabs[i]);
                    _ribbonTabs[i] = null;
                }
            }

            _ribbonTabs = null;

            //Marshal.ReleaseComObject(Globals.InvApp);
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

            var buttonsIdw = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIdw);
            var buttonsIpt = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpt);
            var buttonsIam = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIam);
            var buttonsIpn = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpn);

            if (_defaultButton != null)
            {
                buttonsIdw.CommandControls.AddButton(_defaultButton, true);
                buttonsIpt.CommandControls.AddButton(_defaultButton, true);
                buttonsIam.CommandControls.AddButton(_defaultButton, true);
                buttonsIpn.CommandControls.AddButton(_defaultButton, true);
            }
            if (_info != null)
            {
                infoIdw.CommandControls.AddButton(_info, true);
                infoIpt.CommandControls.AddButton(_info, true);
                infoIam.CommandControls.AddButton(_info, true);
                infoIpn.CommandControls.AddButton(_info, true);
            }
        }

        private void UiEventsOnResetRibbonInterface(NameValueMap context)
        {
	        AddToUserInterface();
        }
    }
}
