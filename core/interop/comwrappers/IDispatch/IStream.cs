#nullable enable

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


/// <summary>
/// COM IStream interface. <see href="https://docs.microsoft.com/en-us/windows/desktop/api/objidl/nn-objidl-istream"/>
/// </summary>
/// <remarks>
/// The definition in <see cref="System.Runtime.InteropServices.ComTypes"/> does not lend
/// itself to efficiently accessing / implementing IStream.
/// </remarks>
[//ComImport,
    Guid("0000000C-0000-0000-C000-000000000046")]//,
                                                 //InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IStream
{
    // pcbRead is optional so it must be a pointer
    unsafe void Read(byte* pv, uint cb, uint* pcbRead);

    // pcbWritten is optional so it must be a pointer
    unsafe void Write(byte* pv, uint cb, uint* pcbWritten);

    // SeekOrgin matches the native values, plibNewPosition is optional
    unsafe void Seek(long dlibMove, SeekOrigin dwOrigin, ulong* plibNewPosition);

    void SetSize(ulong libNewSize);

    // pcbRead and pcbWritten are optional
    unsafe void CopyTo(
        IStream pstm,
        ulong cb,
        ulong* pcbRead,
        ulong* pcbWritten);

    void Commit(uint grfCommitFlags);

    void Revert();

    // Using PreserveSig to allow explicitly returning the HRESULT for "not supported".

    [PreserveSig]
    HRESULT LockRegion(
        ulong libOffset,
        ulong cb,
        uint dwLockType);

    [PreserveSig]
    HRESULT UnlockRegion(
        ulong libOffset,
        ulong cb,
        uint dwLockType);

    void Stat(
        out STATSTG pstatstg,
        STATFLAG grfStatFlag);

    IStream Clone();
}

internal enum HRESULT : int
{
    S_OK = 0,
    S_FALSE = 1,
    E_NOTIMPL = unchecked((int)0x80004001),
    E_ABORT = unchecked((int)0x80004004),
    E_FAIL = unchecked((int)0x80004005),
    E_UNEXPECTED = unchecked((int)0x8000FFFF),
    STG_E_INVALIDFUNCTION = unchecked((int)0x80030001L),
    STG_E_INVALIDPARAMETER = unchecked((int)0x80030057),
    STG_E_INVALIDFLAG = unchecked((int)0x800300FF),
    E_ACCESSDENIED = unchecked((int)0x80070005),
    E_INVALIDARG = unchecked((int)0x80070057),
}
[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
internal struct STATSTG
{
    /// <summary>
    /// Pointer to the name.
    /// </summary>
    private IntPtr pwcsName;
    public STGTY type;

    /// <summary>
    /// Size of the stream in bytes.
    /// </summary>
    public ulong cbSize;

    public FILETIME mtime;
    public FILETIME ctime;
    public FILETIME atime;

    /// <summary>
    /// The stream mode.
    /// </summary>
    public STGM grfMode;

    /// <summary>
    /// Supported locking modes.
    /// <see href="https://docs.microsoft.com/en-us/windows/desktop/api/objidl/ne-objidl-taglocktype"/>
    /// </summary>
    /// <remarks>
    /// '0' means does not support lock modes.
    /// </remarks>
    public uint grfLocksSupported;

    /// <remarks>
    /// Only for IStorage objects
    /// </remarks>
    public Guid clsid;

    /// <remarks>
    /// Only valid for IStorage objects.
    /// </remarks>
    public uint grfStateBits;
    public uint reserved;

    public string? GetName() => Marshal.PtrToStringUni(pwcsName);

    /// <summary>
    /// Caller is responsible for freeing the name memory.
    /// </summary>
    public void FreeName()
    {
        if (pwcsName != IntPtr.Zero)
            Marshal.FreeCoTaskMem(pwcsName);

        pwcsName = IntPtr.Zero;
    }

    /// <summary>
    /// Callee is repsonsible for allocating the name memory.
    /// </summary>
    public void AllocName(string? name)
    {
        pwcsName = Marshal.StringToCoTaskMemUni(name);
    }
}

internal enum STGTY : uint
{
    STGTY_STORAGE = 1,
    STGTY_STREAM = 2,
    STGTY_LOCKBYTES = 3,
    STGTY_PROPERTY = 4
}

[Flags]
internal enum STGM : uint
{
    /// <summary>
    /// Read only, and each change to a storage or stream element is written as it occurs.
    /// Fails if the given storage object already exists.
    /// [STGM_DIRECT] [STGM_READ] [STGM_FAILIFTHERE] [STGM_SHARE_DENY_WRITE]
    /// </summary>
    Default = 0x00000000,

    STGM_TRANSACTED = 0x00010000,
    STGM_SIMPLE = 0x08000000,
    STGM_WRITE = 0x00000001,
    STGM_READWRITE = 0x00000002,
    STGM_SHARE_DENY_NONE = 0x00000040,
    STGM_SHARE_DENY_READ = 0x00000030,
    STGM_SHARE_DENY_WRITE = 0x00000020,
    STGM_SHARE_EXCLUSIVE = 0x00000010,
    STGM_PRIORITY = 0x00040000,
    STGM_DELETEONRELEASE = 0x04000000,
    STGM_NOSCRATCH = 0x00100000,
    STGM_CREATE = 0x00001000,
    STGM_CONVERT = 0x00020000,
    STGM_NOSNAPSHOT = 0x00200000,
    STGM_DIRECT_SWMR = 0x00400000
}

internal enum STATFLAG : uint
{
    /// <summary>
    /// Stat includes the name.
    /// </summary>
    STATFLAG_DEFAULT = 0,

    /// <summary>
    /// Stat doesn't include the name.
    /// </summary>
    STATFLAG_NONAME = 1
}
