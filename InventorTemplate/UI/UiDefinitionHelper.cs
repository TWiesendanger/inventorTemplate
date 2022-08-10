using System;
using System.Windows.Forms;
using Inventor;
using IPictureDisp = stdole.IPictureDisp;

namespace InventorTemplate.UI
{
    public class UiDefinitionHelper
    {
        public static ButtonDefinition CreateButton(string displayText, string internalName, string iconPath, string theme)
        {
            var myButton = new UiButton()
            {
                Bd = CreateButtonDefinition(displayText, internalName, "", iconPath, theme)
            };
            return myButton.Bd;
        }

        public static ButtonDefinition CreateButtonDefinition(string displayName, string internalName,
            string toolTip, string iconFolder, string theme)
        {
            // Check to see if a command already exists is the specified internal name.
            UiDefinitionHelper testDef = null;
            try
            {
                // ReSharper disable once SuspiciousTypeConversion.Global
                testDef = (UiDefinitionHelper)Globals.InvApp.CommandManager.ControlDefinitions[internalName];
            }
            catch
            {
                // checked for null}
            }

            if (testDef != null)
            {
                MessageBox.Show(
                    @"Error when loading the add-in ""InventorTemplate"". A command already exists with the same internal name. Each add-in must have a unique internal name. Change the internal name in the call to CreateButtonDefinition.",
                    @"CSharp Inventor Add-In Template");
                return null;
            }

            iconFolder = GetIconFolder(iconFolder);
            var (iPicDisp16X16, iPicDisp32X32) = GetButtonIcons(iconFolder, theme);

            try
            {
                // Get the ControlDefinitions collection.
                ControlDefinitions controlDefs = Globals.InvApp.CommandManager.ControlDefinitions;

                // Create the command definition.
                ButtonDefinition btnDef = controlDefs.AddButtonDefinition(displayName, internalName,
                    CommandTypesEnum.kShapeEditCmdType, Globals.AddInClientId, "", toolTip, iPicDisp16X16,
                    iPicDisp32X32);
                return btnDef;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public static string GetIconFolder(string iconFolder)
        {
            // Check to see if the provided folder is a full or relative path.
            if (!string.IsNullOrEmpty(iconFolder))
            {
                if (!System.IO.Directory.Exists(iconFolder))
                {
                    // The folder provided doesn't exist, so assume it is a relative path and
                    // build up the full path.
                    string dllPath =
                        System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                    iconFolder = System.IO.Path.Combine(dllPath ?? throw new InvalidOperationException(),
                        iconFolder);
                }
            }

            return iconFolder;
        }
        public static (IPictureDisp iPicDisp16X16, IPictureDisp iPicDisp32X32) GetButtonIcons(string iconFolder, string theme)
        {
            IPictureDisp iPicDisp16X16 = null;
            IPictureDisp iPicDisp32X32 = null;
            if (!string.IsNullOrEmpty(iconFolder))
            {
                if (System.IO.Directory.Exists(iconFolder))
                {
                    string filename16X16 = System.IO.Path.Combine(iconFolder, $"16x16{theme}.png");
                    string filename32X32 = System.IO.Path.Combine(iconFolder, $"32x32{theme}.png");

                    if (System.IO.File.Exists(filename16X16))
                    {
                        try
                        {
                            System.Drawing.Bitmap image16X16 = new System.Drawing.Bitmap(filename16X16);
                            iPicDisp16X16 = ConvertImage.ConvertImageToIPictureDisp(image16X16);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(
                                @"Unable to load the 16x16.png image from """ + iconFolder + @"""." +
                                System.Environment.NewLine + @"No small icon will be used.", @"Error Loading Icon");
                        }
                    }
                    else
                        MessageBox.Show(
                            @"The icon for the small button does not exist: """ + filename16X16 + @"""." +
                            System.Environment.NewLine + @"No small icon will be used.", @"Error Loading Icon");

                    if (System.IO.File.Exists(filename32X32))
                    {
                        try
                        {
                            System.Drawing.Bitmap image32X32 = new System.Drawing.Bitmap(filename32X32);
                            iPicDisp32X32 = ConvertImage.ConvertImageToIPictureDisp(image32X32);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show(
                                @"Unable to load the 32x32.png image from """ + iconFolder + @"""." +
                                System.Environment.NewLine + @"No large icon will be used.", @"Error Loading Icon");
                        }
                    }
                    else
                        MessageBox.Show(
                            @"The icon for the large button does not exist: """ + filename32X32 + @"""." +
                            System.Environment.NewLine + @"No large icon will be used.", @"Error Loading Icon");
                }
            }

            return (iPicDisp16X16, iPicDisp32X32);
        }
        public static RibbonTab SetupTab(string displayName, string internalName, Ribbon invRibbon)
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
		        ribbonTab = invRibbon.RibbonTabs.Add(displayName, internalName, Globals.AddInClientId);
	        }

	        var setupTabRet = ribbonTab;
	        return setupTabRet;
        }
        public static RibbonPanel SetupPanel(string displayName, string internalName, RibbonTab ribbonTab)
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
		        ribbonPanel = ribbonTab.RibbonPanels.Add(displayName, internalName, Globals.AddInClientId);
	        }

	        var setupPanelRet = ribbonPanel;
	        return setupPanelRet;
        }
    }

    ///<summary>
    /// Class used to convert bitmaps and icons between their .Net native types
    /// and an IPictureDisp object which is what the Inventor API requires.
    /// </summary>
    [System.Security.Permissions.PermissionSet
        (System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public class ConvertImage : AxHost
    {
        public ConvertImage()
            : base("59EE46BA-677D-4d20-BF10-8D8067CB8B32")
        {
        }

        public static stdole.IPictureDisp ConvertImageToIPictureDisp(System.Drawing.Image image)
        {
            try
            {
                return (stdole.IPictureDisp)GetIPictureFromPicture(image);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}