using System;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Inventor;
using InventorTemplate.Helper;
using InventorTemplate.UI.Dialog;

namespace InventorTemplate.UI
{
    public class UiButton
    {
        private ButtonDefinition _bd;

        public ButtonDefinition Bd
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get => _bd;

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_bd != null)
                {
                    _bd.OnExecute -= ButtonOnExecute;
                }

                _bd = value;
                if (_bd != null)
                {
                    _bd.OnExecute += ButtonOnExecute;
                }
            }
        }

        private void ButtonOnExecute(NameValueMap context)
        {
            switch (Bd.InternalName)
            {
                case "DefaultButton":
                    MessageBox.Show(@"Default message.", @"Default title");
                    return;
                case "Info":
                    var infoDlg = new FrmInfo();
                    infoDlg.ShowDialog(new WindowWrapper((IntPtr)Globals.InvApp.MainFrameHWND));
                    return;
                default:
                    return;
            }
        }
    }
}
