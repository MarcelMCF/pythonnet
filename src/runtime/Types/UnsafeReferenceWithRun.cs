using System;
using System.ComponentModel;

namespace Python.Runtime;

[EditorBrowsable(EditorBrowsableState.Never)]
[Obsolete(Util.InternalUseOnly)]
public struct UnsafeReferenceWithRun
{
    internal UnsafeReferenceWithRun(BorrowedReference pyObj)
    {
        RawObj = pyObj.DangerousGetAddressOrNull();
        Run = Runtime.GetRun();
    }

    internal IntPtr RawObj;
    internal BorrowedReference Ref => new(RawObj);
    internal int Run;

    internal BorrowedReference CheckRun()
    {
        if (RawObj == IntPtr.Zero)
            throw new InvalidOperationException(
                "Python object reference is null. The derived object was never properly initialized via Python.");

        if (Run != Runtime.GetRun())
            throw new RuntimeShutdownException(RawObj);

        return Ref;
    }
}
