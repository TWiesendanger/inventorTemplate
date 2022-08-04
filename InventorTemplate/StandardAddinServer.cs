using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.Helper;
using InventorTemplate.UI;

namespace InventorTemplate
{
    [ProgId("InventorTemplate.StandardAddInServer")]
    [Guid(Globals.AddInClientId)]
    public class StandardAddInServer : ApplicationAddInServer
    {

        public delegate ButtonDefinition CreateButton(string displayText, string internalName, string iconPath);
        public ButtonDefinition button_template(string displayText, string internalName, string iconPath)
        {
            var myButton = new UiButton()
            {
                Bd = ButtonDefinitionHelper.CreateButtonDefinition(displayText, internalName, "", iconPath)
            };
            return myButton.Bd;
        }

        // Declare all buttons here
        private ButtonDefinition _createCustomTable;
        private ButtonDefinition _info;
        

        // This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        // to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        // the first time. However, with the introduction of the ribbon this argument is always true.
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                // Initialize AddIn members.
                Globals.InvApp = addInSiteObject.Application;

                // Connect to the user-interface events to handle a ribbon reset.
                MUiEvents = Globals.InvApp.UserInterfaceManager.UserInterfaceEvents;

                _sessionStateObject = new SessionStateObject();
                _log = _sessionStateObject.AsISessionState().LogMan.GetCurrentClassLogger();
                _log.Debug("Addin Evatec Toolbox Activated");

                //_sessionStateObject.AsISessionState().ArgumentChecker.IsNotNull(()=>_log);

                // *********************************************************************************
                // * The remaining code in this Sub is all for adding the add-in into Inventor's UI.
                // * It can be deleted if this add-in doesn't have a UI and only runs in the 
                // * background handling events.
                // *********************************************************************************
                // ButtonName = create_button(display_text, internal_name, icon_path)
                CreateButton createButton = button_template;
                string customTableText = "Prüftabelle erstellen / @aktualisieren";
                customTableText = customTableText.Replace("@", System.Environment.NewLine);
                _createCustomTable = createButton(customTableText, "CreateCustomTable", @"ButtonResources\CreateInspectionTableIcon",
                    _sessionStateObject.AsISessionState());
                _searchDimension = createButton("Bemassung suchen", "SearchDimension", @"ButtonResources\SearchIcon",
                    _sessionStateObject.AsISessionState());
                _updateInspectionDimension = createButton("Prüfmasse aktualisieren", "UpdateInspectionDimension",
                    @"ButtonResources\UpdateInspectionDimIcon", _sessionStateObject.AsISessionState());

                _info = createButton("Info", "Info", @"ButtonResources\InfoIcon", _sessionStateObject.AsISessionState());
                _helpfile = createButton("Hilfe", "Help", @"ButtonResources\HelpIcon", _sessionStateObject.AsISessionState());
                _reposition = createButton("Neu positionieren", "Reposition", @"ButtonResources\UpdateInspectionDimIcon",
                    _sessionStateObject.AsISessionState());
                _savePosition = createButton("Positionen speichern", "SavePosition", @"ButtonResources\SavePositionIcon",
                    _sessionStateObject.AsISessionState());


                // ExternalUtilitiesDrawing
                _drawingResCleanUp = createButton("Zeich Res bereinigen", "DrawingResCleanUp", @"ButtonResources\DrawingResCleanUp",
                    _sessionStateObject.AsISessionState());
                _resetBrowserNames = createButton("Browsernamen", "BrowserName", @"ButtonResources\BrowserName",
                    _sessionStateObject.AsISessionState());
                _resetBrowserNamesIam = createButton("Browsernamen", "BrowserNameIam", @"ButtonResources\BrowserName",
                    _sessionStateObject.AsISessionState());
                _pdfInternalExport = createButton("PDFExportInternal", "PdfExportInternal",
                    @"ButtonResources\PdfExportInternal", _sessionStateObject.AsISessionState());
                _drawingEdits = createButton("Zeichnung bearbeiten", "DrawingEdit", @"ButtonResources\DrawingEdit",
                    _sessionStateObject.AsISessionState());
                _zUpdate = createButton("ZUpdate", "zUpdate", @"ButtonResources\ZUpdate",
                    _sessionStateObject.AsISessionState());
                _checkVersionTable = createButton("Version Prüfmasstabelle", "checkVersionTable", @"ButtonResources\CheckVersionTable",
                    _sessionStateObject.AsISessionState());

                // ExternalUtilitiesModels
                _mUpdate = createButton("MUpdate", "mUpdate", @"ButtonResources\MUpdate",
                    _sessionStateObject.AsISessionState());
                _preview = createButton("Preview", "preview", @"ButtonResources\Preview",
                    _sessionStateObject.AsISessionState());
                _sanityCheck = createButton("Model Check", "sanityCheck", @"ButtonResources\SanityCheck",
                    _sessionStateObject.AsISessionState());

                // ExternalUtilitiesIPN
                _syncToIpn = createButton("SyncToIPN", "syncToIpn", @"ButtonResources\SyncToIpn",
                    _sessionStateObject.AsISessionState());

                // AdminPanel
                _deleteNonVisibleHoleThreadNotes = createButton("DeleteNonVisibleHoleThreadNotes", "deleteNonVisibleHoleThreadNotes", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());
                _deleteHoleThreadNotes = createButton("deleteHoleThreadNotes", "deleteHoleThreadNotes", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());
                _deleteChamferNotes = createButton("deleteChamferNotes", "deleteChamferNotes", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());
                _deleteSurfaceTextureSymbols = createButton("deleteSurfaceTextureSymbols", "deleteSurfaceTextureSymbols", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());
                _deleteDrawingDimensions = createButton("deleteDrawingDimensions", "deleteDrawingDimensions", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());
                _deleteFeatureControlFrames = createButton("deleteFeatureControlFrames", "deleteFeatureControlFrames", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());

                _collapseNodes = createButton("collapseNodes", "collapseNodes", @"ButtonResources\AdminPanel",
                    _sessionStateObject.AsISessionState());

                // Add to the user interface, if it's the first time.
                // If this add-in doesn't have a UI but runs in the background listening
                // to events, you can delete this.
                if (firstTime)
                    AddToUserInterface();
            }
            catch (Exception ex)
            {
                MessageBox.Show(@"Unexpected failure in the activation of the add-in ""Evatec_Toolbox""" + System.Environment.NewLine + System.Environment.NewLine + ex.Message);
            }
        }

        // This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        // unloaded either manually by the user or when the Inventor session is terminated.
        public void Deactivate()
        {
            // Release objects.
            _createCustomTable = null;
            _info = null;
            MUiEvents = null;
            Globals.InvApp = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // This property is provided to allow the AddIn to expose an API of its own to other 
        // programs. Typically, this  would be done by implementing the AddIn's API
        // interface in a class and returning that class object through this property.
        // Typically it's not used, like in this case, and returns Nothing.
        public object Automation => null;

        // Note:this method is now obsolete, you should use the 
        // ControlDefinition functionality for implementing commands.
        public void ExecuteCommand(int commandId)
        { }

        // Adds whatever is needed by this add-in to the user-interface.  This is 
        // called when the add-in loaded and also if the user interface is reset.
        private void AddToUserInterface()
        {

            var iamRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Assembly"];
            var iptRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Part"];
            var idwRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Drawing"];
            var ipnRibbon = Globals.InvApp.UserInterfaceManager.Ribbons["Presentation"];


            var tabIdw = SetupTab("Evatec_Toolbox", "Evatec_Toolbox", idwRibbon);
            var tabIpt = SetupTab("Evatec_Toolbox", "Evatec_Toolbox", iptRibbon);
            var tabIam = SetupTab("Evatec_Toolbox", "Evatec_Toolbox", iamRibbon);
            var tabIpn = SetupTab("Evatec_Toolbox", "Evatec_Toolbox", ipnRibbon);



            var inspectionDimensionPan = SetupPanel("Prüfbemassung", "Pruefbemassung", tabIdw);
            var externalUtilitiesPanIdw = SetupPanel("Utilities", "ExternalUtilitiesIDW", tabIdw);
            var externalUtilitiesPanIpt = SetupPanel("Utilities", "ExternalUtilitiesIPT", tabIpt);
            var externalUtilitiesPanIam = SetupPanel("Utilities", "ExternalUtilitiesIAM", tabIam);
            var externalUtilitiesPanIpn = SetupPanel("Utilities", "ExternalUtilitiesIAM", tabIpn);
            var adminIdw = SetupPanel("Admin", "adminIdw", tabIpn);
            var bomCheckPanIam = SetupPanel("BomCheck", "bomCheckIAM", tabIam);

            var infoIdw = SetupPanel("Info", "Info", tabIdw);
            var infoIpt = SetupPanel("Info", "Info", tabIpt);
            var infoIam = SetupPanel("Info", "Info", tabIam);
            var infoIpn = SetupPanel("Info", "Info", tabIpn);

            // Set up Buttons.
            if (_createCustomTable != null) inspectionDimensionPan.CommandControls.AddButton(_createCustomTable, true);
            if (_searchDimension != null) inspectionDimensionPan.CommandControls.AddButton(_searchDimension, true);
            if (_updateInspectionDimension != null) inspectionDimensionPan.CommandControls.AddButton(_updateInspectionDimension, true);
            if (_info != null)
            {
                infoIdw.CommandControls.AddButton(_info, true);
                infoIpt.CommandControls.AddButton(_info, true);
                infoIam.CommandControls.AddButton(_info, true);
                infoIpn.CommandControls.AddButton(_info, true);
            }

            if (_helpfile != null)
            {
                infoIdw.CommandControls.AddButton(_helpfile, true);
                infoIpt.CommandControls.AddButton(_helpfile, true);
                infoIam.CommandControls.AddButton(_helpfile, true);
                infoIpn.CommandControls.AddButton(_helpfile, true);
            }
            if (_reposition != null) inspectionDimensionPan.CommandControls.AddButton(_reposition, true);
            if (_savePosition != null) inspectionDimensionPan.CommandControls.AddButton(_savePosition, true);

            if (_zUpdate != null) externalUtilitiesPanIdw.CommandControls.AddButton(_zUpdate, true);
            if (_drawingEdits != null) externalUtilitiesPanIdw.CommandControls.AddButton(_drawingEdits, true);
            if (_pdfInternalExport != null) externalUtilitiesPanIdw.CommandControls.AddButton(_pdfInternalExport, true);
            if (_resetBrowserNamesIam != null) externalUtilitiesPanIam.CommandControls.AddButton(_resetBrowserNamesIam, true);
            if (_resetBrowserNames != null)
            {
                externalUtilitiesPanIdw.CommandControls.AddButton(_resetBrowserNames, true);
                externalUtilitiesPanIpt.CommandControls.AddButton(_resetBrowserNames, true);
            }
            if (_drawingResCleanUp != null) externalUtilitiesPanIdw.CommandControls.AddButton(_drawingResCleanUp, true);
            if (_mUpdate != null)
            {
                externalUtilitiesPanIam.CommandControls.AddButton(_mUpdate, true);
                externalUtilitiesPanIpt.CommandControls.AddButton(_mUpdate, true);
            }

            if (_preview != null)
            {
                externalUtilitiesPanIam.CommandControls.AddButton(_preview, true);
                externalUtilitiesPanIpt.CommandControls.AddButton(_preview, true);
            }
            if (_sanityCheck != null)
            {
                externalUtilitiesPanIam.CommandControls.AddButton(_sanityCheck, true);
                externalUtilitiesPanIpt.CommandControls.AddButton(_sanityCheck, true);
            }

            if (_syncToIpn != null) externalUtilitiesPanIpn.CommandControls.AddButton(_syncToIpn, true);
            if (_checkVersionTable != null) externalUtilitiesPanIdw.CommandControls.AddButton(_checkVersionTable, true);

            if (Settings.Default.AdminPanel)
            {
                if (_deleteNonVisibleHoleThreadNotes != null)
                    adminIdw.CommandControls.AddButton(_deleteNonVisibleHoleThreadNotes, true);
                if (_deleteHoleThreadNotes != null)
                    adminIdw.CommandControls.AddButton(_deleteHoleThreadNotes, true);
                if (_deleteChamferNotes != null)
                    adminIdw.CommandControls.AddButton(_deleteChamferNotes, true);
                if (_deleteSurfaceTextureSymbols != null)
                    adminIdw.CommandControls.AddButton(_deleteSurfaceTextureSymbols, true);
                if (_deleteDrawingDimensions != null)
                    adminIdw.CommandControls.AddButton(_deleteDrawingDimensions, true);
                if (_deleteFeatureControlFrames != null)
                    adminIdw.CommandControls.AddButton(_deleteFeatureControlFrames, true);
            }

            if (_collapseNodes != null)
                adminIdw.CommandControls.AddButton(_collapseNodes, true);
        }

        private static RibbonTab SetupTab(string displayName, string internalName, Ribbon invRibbon)
        {
            RibbonTab ribbonTab = null;
            try
            {
                ribbonTab = invRibbon.RibbonTabs[internalName];
            }
            catch (Exception)
            {
                // ignored
            }

            if (ribbonTab == null)
            {
                ribbonTab = invRibbon.RibbonTabs.Add(displayName, internalName, Globals.GAddInClientId);
            }

            var setupTabRet = ribbonTab;
            return setupTabRet;
        }

        private static RibbonPanel SetupPanel(string displayName, string internalName, RibbonTab ribbonTab)
        {
            RibbonPanel ribbonPanel = null;
            try
            {
                ribbonPanel = ribbonTab.RibbonPanels[internalName];
            }
            catch (Exception)
            {
                // ignored
            }

            if (ribbonPanel == null)
            {
                ribbonPanel = ribbonTab.RibbonPanels.Add(displayName, internalName, Globals.GAddInClientId);
            }

            var setupPanelRet = ribbonPanel;
            return setupPanelRet;
        }

        private void m_uiEvents_OnResetRibbonInterface(NameValueMap context)
        {
            // The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface();
        }
    }
}
