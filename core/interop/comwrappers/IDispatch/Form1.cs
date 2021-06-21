using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Internal;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ComWrappersIDispatch
{
    public partial class Form1 : Form
    {
        public unsafe Form1()
        {
            InitializeComponent();

            this.webBrowser1.DocumentText =
$@"<button onClick=""dispatch()"">Send to .NET!</button>
<input type=""text"" id=""message"" value=""From the Web Browser!""></input>

<script>
    function dispatch() {{
        window.external.Func1();
        window.external.Func2(27);
        window.external.Func3(document.getElementById(""message"").value);
    }}
</script>";

            this.Controls.Add(this.webBrowser1);

            Icon i = SystemIcons.Hand;
            PICTDESC pictdesc = PICTDESC.CreateIconPICTDESC(i.Handle);
            Guid g = IPicture.Guid;
            IntPtr lpPicture;
            HRESULT result = OleCreatePictureIndirect(pictdesc, &g, false, &lpPicture);
            if (result != HRESULT.S_OK)
            {
                throw new COMException(null, (int)result);
            }

            IPicture picture = (IPicture)ComWrappersImpl.Instance.GetOrCreateObjectForComInstance(lpPicture, CreateObjectFlags.None);

            using var memStream = new MemoryStream();
            var gpStream = new GPStream(memStream, makeSeekable: false);
            IntPtr streamPtr = ComWrappersImpl.Instance.GetOrCreateComInterfaceForObject(gpStream, CreateComInterfaceFlags.None);

            var inst = Marshal.PtrToStructure<ComWrappersImpl.VtblPtr>(streamPtr);
            Debug.WriteLine($"vtbl: {inst.Vtbl}");

            picture.SaveAsFile(streamPtr, -1, out _);

            Marshal.Release(streamPtr);
            Marshal.Release(lpPicture);
        }

        [DllImport("oleaut32.dll")]
        internal static unsafe extern HRESULT OleCreatePictureIndirect(PICTDESC pictdesc, Guid* refiid, bool fOwn, IntPtr* lplpvObj);

        private void AttachObject(object obj)
        {
            var proxy = new AnyObjectProxy(obj);
            this.webBrowser1.ObjectForScripting = proxy;
        }

        private class ExposedObject
        {
            private readonly Form1 form;
            public ExposedObject(Form1 form)
            {
                this.form = form;
            }
            public void Func1()
            {
                Debug.WriteLine($"{nameof(Func1)}");
            }
            public void Func2(int a)
            {
                Debug.WriteLine($"{nameof(Func2)}({a})");
            }
            public void Func3(string msg)
            {
                Debug.WriteLine($"{nameof(Func3)}({msg})");
                this.form.textBox1.Text = msg;
            }
        }
    }

    internal static class Ole
    {
        public const int PICTYPE_ICON = 3;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal sealed class PICTDESC
    {
        internal int cbSizeOfStruct;
        public int picType;
        internal IntPtr union1;
        internal int union2;
        internal int union3;

        public static PICTDESC CreateIconPICTDESC(IntPtr hicon)
        {
            return new PICTDESC()
            {
                cbSizeOfStruct = 12,
                picType = Ole.PICTYPE_ICON,
                union1 = hicon
            };
        }
    }
}
