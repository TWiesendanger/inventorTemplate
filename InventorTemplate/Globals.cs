using Inventor;

namespace InventorTemplate
{
    public static class Globals
    {
        /// <summary>Holds a reference to the app (inventor)</summary>
        public static Application InvApp;
        /// <summary>Holds a reference to the addin.</summary>
        public static ApplicationAddInSite InvApplicationAddInSite;
        /// <summary>This is the guid used through the whole addin. This was generated during the creation of the addin. Make sure it is unique. There cannot be another addin having the same guid.</summary>
        public const string AddInClientId = "74F1F066-BB65-4141-BDC8-C4A7669913F1";
        /// <summary>Holds a reference to the active app them. Can be darktheme or lighttheme.</summary>
        public static Theme ActiveTheme;
    }
}