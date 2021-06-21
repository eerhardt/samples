#nullable enable

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


//[ComImport]
//[Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")]
//[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IPicture
{
    static readonly Guid Guid = new Guid(0x7BF80980, 0xBF32, 0x101A, 0x8B, 0xBB, 0, 0xAA, 0x00, 0x30, 0x0C, 0xAB);

    IntPtr GetHandle();

    IntPtr GetHPal();

    [return: MarshalAs(UnmanagedType.I2)]
    short GetPictureType();

    int GetWidth();

    int GetHeight();

    void Render();

    void SetHPal([In] IntPtr phpal);

    IntPtr GetCurDC();

    void SelectPicture([In] IntPtr hdcIn,
                       [Out, MarshalAs(UnmanagedType.LPArray)] int[] phdcOut,
                       [Out, MarshalAs(UnmanagedType.LPArray)] int[] phbmpOut);

    [return: MarshalAs(UnmanagedType.Bool)]
    bool GetKeepOriginalFormat();

    void SetKeepOriginalFormat([In, MarshalAs(UnmanagedType.Bool)] bool pfkeep);

    void PictureChanged();

    int SaveAsFile(IntPtr pstm,
                   int fSaveMemCopy,
                   out int pcbSize);

    int GetAttributes();

    void SetHdc([In] IntPtr hdc);
}
