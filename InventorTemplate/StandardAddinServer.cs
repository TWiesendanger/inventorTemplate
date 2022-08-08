using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.Helper;
using InventorTemplate.Helper.Logging;
using Microsoft.Extensions.Configuration;
using NLog;
using NLog.Extensions.Logging;

namespace InventorTemplate
{
    [ProgId("InventorTemplate.StandardAddInServer")]
    [Guid(Globals.AddInClientId)]
    public class StandardAddInServer : ApplicationAddInServer
    {
	    private UserInterfaceEvents _uiEvents;
	    private ButtonDefinition _button1;
	    private ButtonDefinition _info;
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
	        MessageBox.Show(@"Test");

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false);

            IConfiguration config = builder.Build();
            LogManager.Configuration = new NLogLoggingConfiguration(config.GetSection("NLog"));


            //var logSettings = config.GetSection("Logging").Get<Logging>();
            //LogManager.Configuration.Variables["logPath"] = logSettings.LogPath;
            var logger = LogManager.GetCurrentClassLogger();

            try
            {
                logger.Debug("Addin InventorTemplate Activated");

                UiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;

                _info = UiDefinitionHelper.CreateButton("Info", "Info", @"UI\ButtonResources\InfoIcon");
                _button1 = UiDefinitionHelper.CreateButton("Button1", "button1", @"UI\ButtonResources\Button1");

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
	        _button1 = null;
            _info = null;
            Globals.InvApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public object Automation => null;
        public void ExecuteCommand(int commandId) { }

        private void AddToUserInterface()
        {
	        var iamRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            var iptRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            var idwRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            var ipnRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Presentation"];

            var tabIdw = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", idwRibbon);
            var tabIpt = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", iptRibbon);
            var tabIam = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", iamRibbon);
            var tabIpn = UiDefinitionHelper.SetupTab("InventorTemplate", "InventorTemplate", ipnRibbon);

            var infoIdw = UiDefinitionHelper.SetupPanel("Info", "Info", tabIdw);
            var infoIpt = UiDefinitionHelper.SetupPanel("Info", "Info", tabIpt);
            var infoIam = UiDefinitionHelper.SetupPanel("Info", "Info", tabIam);
            var infoIpn = UiDefinitionHelper.SetupPanel("Info", "Info", tabIpn);

            var buttonsIdw = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIdw);
            var buttonsIpt = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpt);
            var buttonsIam = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIam);
            var buttonsIpn = UiDefinitionHelper.SetupPanel("AddinCommands", "AddinCommands", tabIpn);

            if (_button1 != null)
            {
	            buttonsIdw.CommandControls.AddButton(_button1, true);
	            buttonsIpt.CommandControls.AddButton(_button1, true);
	            buttonsIam.CommandControls.AddButton(_button1, true);
	            buttonsIpn.CommandControls.AddButton(_button1, true);
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
