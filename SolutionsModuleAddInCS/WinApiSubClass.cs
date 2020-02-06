using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolutionsModuleAddInCS
{
    class WinApiSubClass : NativeWindow
    {
        public delegate void CallbackProcHandler(ref Message Msg);
        public event CallbackProcHandler CallbackProc;
        public bool Subclassed { get; set; } = false;

        public WinApiSubClass(IntPtr handle)
        {
            base.AssignHandle(handle);
            if (handle != IntPtr.Zero)
                Subclassed = true;
        }

        protected override void WndProc(ref Message m)
        {
            if (Subclassed)
                CallbackProc?.Invoke(ref m);

            base.WndProc(ref m);
        }

        public void Release()
        {
            base.ReleaseHandle();
        }
    }
}
