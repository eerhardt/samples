#nullable enable

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ComTypes = System.Runtime.InteropServices.ComTypes;

internal unsafe class ComWrappersImpl : ComWrappers
{
    private static readonly ComInterfaceEntry* wrapperEntry;// = InitializeComInterfaceEntry();
    private static readonly Lazy<ComWrappersImpl> g_ComWrappers = new Lazy<ComWrappersImpl>(() => new ComWrappersImpl(), true);

    internal static ComWrappersImpl Instance => g_ComWrappers.Value;

    //private static ComInterfaceEntry* InitializeComInterfaceEntry()
    static ComWrappersImpl()
    {
        GetIUnknownImpl(out IntPtr fpQueryInteface, out IntPtr fpAddRef, out IntPtr fpRelease);

        var vtbl = new IStreamVtbl()
        {
            IUnknownImpl = new IUnknownVtbl()
            {
                QueryInterface = fpQueryInteface,
                AddRef = fpAddRef,
                Release = fpRelease
            },
            Read = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pRead),
            Write = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pWrite),
            Seek = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pSeek),
            SetSize = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pSetSize),
            CopyTo = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pCopyTo),
            Commit = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pCommit),
            Revert = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pRevert),
            LockRegion = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pLockRegion),
            UnlockRegion = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pUnlockRegion),
            Stat = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pStat),
            Clone = Marshal.GetFunctionPointerForDelegate(IStreamVtbl.pClone),
        };
        var vtblRaw = RuntimeHelpers.AllocateTypeAssociatedMemory(typeof(IStreamVtbl), sizeof(IStreamVtbl));
        Marshal.StructureToPtr(vtbl, vtblRaw, false);

        wrapperEntry = (ComInterfaceEntry*)RuntimeHelpers.AllocateTypeAssociatedMemory(typeof(IStreamVtbl), sizeof(ComInterfaceEntry));
        wrapperEntry->IID = new Guid("0000000C-0000-0000-C000-000000000046");
        wrapperEntry->Vtable = vtblRaw;

        Debug.WriteLine($"vtable {vtblRaw}");
        Debug.WriteLine($"vtable {wrapperEntry->Vtable}");
        //return wrapperEntry;
    }

    protected override unsafe ComInterfaceEntry* ComputeVtables(object obj, CreateComInterfaceFlags flags, out int count)
    {
        Debug.Assert(obj is IStream);
        Debug.Assert(wrapperEntry != null);
        Debug.WriteLine($"IID {wrapperEntry->IID}");
        Debug.WriteLine($"vtable {wrapperEntry->Vtable}");

        // Always return the same table mappings.
        count = 1;
        return wrapperEntry;
    }

    protected override object CreateObject(IntPtr externalComObject, CreateObjectFlags flags)
    {
        Debug.Assert(flags == CreateObjectFlags.None);

        Guid pictureGuid = IPicture.Guid;
        IntPtr comObject;
        int hr = Marshal.QueryInterface(externalComObject, ref pictureGuid, out comObject);
        if (hr == 0)
            return new PictureWrapper(comObject);

        throw new NotImplementedException();
    }

    protected override void ReleaseObjects(IEnumerable objects)
    {
        throw new NotImplementedException();
    }

    public struct IUnknownVtbl
    {
        public IntPtr QueryInterface;
        public IntPtr AddRef;
        public IntPtr Release;
    }

    public struct IStreamVtbl
    {
        public IUnknownVtbl IUnknownImpl;
        public IntPtr Read;
        public IntPtr Write;
        public IntPtr Seek;
        public IntPtr SetSize;
        public IntPtr CopyTo;
        public IntPtr Commit;
        public IntPtr Revert;
        public IntPtr LockRegion;
        public IntPtr UnlockRegion;
        public IntPtr Stat;
        public IntPtr Clone;

        public delegate void _Read(IntPtr thisPtr, byte* pv, uint cb, uint* pcbRead);
        public delegate void _Write(IntPtr thisPtr, byte* pv, uint cb, uint* pcbWritten);
        public delegate void _Seek(IntPtr thisPtr, long dlibMove, SeekOrigin dwOrigin, ulong* plibNewPosition);
        public delegate void _SetSize(IntPtr thisPtr, ulong libNewSize);
        public delegate void _CopyTo(
            IntPtr thisPtr,
            IntPtr pstm,
            ulong cb,
            ulong* pcbRead,
            ulong* pcbWritten);
        public delegate void _Commit(IntPtr thisPtr, uint grfCommitFlags);
        public delegate void _Revert(IntPtr thisPtr);
        public delegate HRESULT _LockRegion(
            IntPtr thisPtr,
            ulong libOffset,
            ulong cb,
            uint dwLockType);
        public delegate HRESULT _UnlockRegion(
            IntPtr thisPtr,
            ulong libOffset,
            ulong cb,
            uint dwLockType);
        public delegate void _Stat(
            IntPtr thisPtr,
            out STATSTG pstatstg,
            STATFLAG grfStatFlag);
        public delegate IntPtr _Clone(IntPtr thisPtr);

        public static _Read pRead = new _Read(ReadInternal);
        public static _Write pWrite = new _Write(WriteInternal);
        public static _Seek pSeek = new _Seek(SeekInternal);
        public static _SetSize pSetSize = new _SetSize(SetSizeInternal);
        public static _CopyTo pCopyTo = new _CopyTo(CopyToInternal);
        public static _Commit pCommit = new _Commit(CommitInternal);
        public static _Revert pRevert = new _Revert(RevertInternal);
        public static _LockRegion pLockRegion = new _LockRegion(LockRegionInternal);
        public static _UnlockRegion pUnlockRegion = new _UnlockRegion(UnlockRegionInternal);
        public static _Stat pStat = new _Stat(StatInternal);
        public static _Clone pClone = new _Clone(CloneInternal);

        public static void ReadInternal(IntPtr thisPtr, byte* pv, uint cb, uint* pcbRead)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Read(pv, cb, pcbRead);
        }

        public static void WriteInternal(IntPtr thisPtr, byte* pv, uint cb, uint* pcbWritten)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Write(pv, cb, pcbWritten);
        }

        public static void SeekInternal(IntPtr thisPtr, long dlibMove, SeekOrigin dwOrigin, ulong* plibNewPosition)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Seek(dlibMove, dwOrigin, plibNewPosition);
        }

        public static void SetSizeInternal(IntPtr thisPtr, ulong libNewSize)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.SetSize(libNewSize);
        }

        public static void CopyToInternal(
            IntPtr thisPtr,
            IntPtr pstm,
            ulong cb,
            ulong* pcbRead,
            ulong* pcbWritten)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            var pstmStream = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)pstm);
            inst.CopyTo(pstmStream, cb, pcbRead, pcbWritten);
        }

        public static void CommitInternal(IntPtr thisPtr, uint grfCommitFlags)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Commit(grfCommitFlags);
        }

        public static void RevertInternal(IntPtr thisPtr)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Revert();
        }

        public static HRESULT LockRegionInternal(IntPtr thisPtr, ulong libOffset, ulong cb, uint dwLockType)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            return inst.LockRegion(libOffset, cb, dwLockType);
        }

        public static HRESULT UnlockRegionInternal(IntPtr thisPtr, ulong libOffset, ulong cb, uint dwLockType)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            return inst.UnlockRegion(libOffset, cb, dwLockType);
        }

        public static void StatInternal(IntPtr thisPtr, out STATSTG pstatstg, STATFLAG grfStatFlag)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);
            inst.Stat(out pstatstg, grfStatFlag);
        }

        public static IntPtr CloneInternal(IntPtr thisPtr)
        {
            var inst = ComInterfaceDispatch.GetInstance<IStream>((ComInterfaceDispatch*)thisPtr);

            return Instance.GetOrCreateComInterfaceForObject(inst.Clone(), CreateComInterfaceFlags.None);
        }
    }

    public struct IPictureVtbl
    {
        internal delegate int _SaveAsFile(IntPtr thisPtr, IntPtr pstm, int fSaveMemCopy, out int pcbSize);

        public IUnknownVtbl IUnknownImpl;
        public IntPtr GetHandle;
        public IntPtr GetHPal;
        public IntPtr GetPictureType;
        public IntPtr GetWidth;
        public IntPtr GetHeight;
        public IntPtr Render;
        public IntPtr SetHPal;
        public IntPtr GetCurDC;
        public IntPtr SelectPicture;
        public IntPtr GetKeepOriginalFormat;
        public IntPtr SetKeepOriginalFormat;
        public IntPtr PictureChanged;
        public _SaveAsFile SaveAsFile;
        public IntPtr GetAttributes;
        public IntPtr SetHdc;
    }

    public struct VtblPtr
    {
        public IntPtr Vtbl;
    }

    private class PictureWrapper : IPicture
    {
        private readonly IntPtr _wrappedInstance;
        private readonly IPictureVtbl _vtable;

        public PictureWrapper(IntPtr wrappedInstance)
        {
            _wrappedInstance = wrappedInstance;

            var inst = Marshal.PtrToStructure<VtblPtr>(_wrappedInstance);
            _vtable = Marshal.PtrToStructure<IPictureVtbl>(inst.Vtbl);
        }

        public int SaveAsFile(IntPtr pstm, int fSaveMemCopy, out int pcbSize)
        {
            var inst = Marshal.PtrToStructure<VtblPtr>(pstm);
            Debug.WriteLine($"vtbl: {inst.Vtbl}");

            return _vtable.SaveAsFile(_wrappedInstance, pstm, fSaveMemCopy, out pcbSize);
        }

        // The following are not implemented because they are never invoked
        public int GetAttributes() => throw new NotImplementedException();
        public IntPtr GetCurDC() => throw new NotImplementedException();
        public IntPtr GetHandle() => throw new NotImplementedException();
        public int GetHeight() => throw new NotImplementedException();
        public IntPtr GetHPal() => throw new NotImplementedException();
        [return: MarshalAs(UnmanagedType.Bool)]
        public bool GetKeepOriginalFormat() => throw new NotImplementedException();
        [return: MarshalAs(UnmanagedType.I2)]
        public short GetPictureType() => throw new NotImplementedException();
        public int GetWidth() => throw new NotImplementedException();
        public void PictureChanged() => throw new NotImplementedException();
        public void Render() => throw new NotImplementedException();
        public void SelectPicture([In] IntPtr hdcIn, [MarshalAs(UnmanagedType.LPArray), Out] int[] phdcOut, [MarshalAs(UnmanagedType.LPArray), Out] int[] phbmpOut) => throw new NotImplementedException();
        public void SetHdc([In] IntPtr hdc) => throw new NotImplementedException();
        public void SetHPal([In] IntPtr phpal) => throw new NotImplementedException();
        public void SetKeepOriginalFormat([In, MarshalAs(UnmanagedType.Bool)] bool pfkeep) => throw new NotImplementedException();
    }
}
