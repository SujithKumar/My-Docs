using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace eAllocation
{
    public class ListViewHelper
    {
        private ListViewHelper()
        {
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr handle, int messg, int wparam, int lparam);

        public static void SetExtendedStyle(Control control, eAllocation.frmAllocation.ListViewExtendedStyles exStyle)
        {
            eAllocation.frmAllocation.ListViewExtendedStyles styles;
            styles = (eAllocation.frmAllocation.ListViewExtendedStyles)SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.GetExtendedStyle, 0, 0);
            styles |= exStyle;
            SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.SetExtendedStyle, 0, (int)styles);
        }

        public static void EnableDoubleBuffer(Control control)
        {
            eAllocation.frmAllocation.ListViewExtendedStyles styles;
            // read current style
            styles = (eAllocation.frmAllocation.ListViewExtendedStyles)SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.GetExtendedStyle, 0, 0);
            // enable double buffer and border select
            styles |= eAllocation.frmAllocation.ListViewExtendedStyles.DoubleBuffer | eAllocation.frmAllocation.ListViewExtendedStyles.BorderSelect;
            // write new style
            SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.SetExtendedStyle, 0, (int)styles);
        }
        public static void DisableDoubleBuffer(Control control)
        {
            eAllocation.frmAllocation.ListViewExtendedStyles styles;
            // read current style
            styles = (eAllocation.frmAllocation.ListViewExtendedStyles)SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.GetExtendedStyle, 0, 0);
            // disable double buffer and border select
            styles -= styles & eAllocation.frmAllocation.ListViewExtendedStyles.DoubleBuffer;
            styles -= styles & eAllocation.frmAllocation.ListViewExtendedStyles.BorderSelect;
            // write new style
            SendMessage(control.Handle, (int)eAllocation.frmAllocation.ListViewMessages.SetExtendedStyle, 0, (int)styles);
        }
    }
}
