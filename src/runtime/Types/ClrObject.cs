using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace Python.Runtime
{
    [Serializable]
    [DebuggerDisplay("clrO: {inst}")]
    internal sealed class CLRObject : ManagedType
    {
        internal readonly object inst;

        // "borrowed" references — tracks all Python wrappers for CLR objects.
        // Entries are added in Create/OnLoad and removed by:
        //   • ClassBase.tp_clear        (regular CLR wrappers, called from tp_dealloc)
        //   • ClassDerivedObject.tp_dealloc (Python-derived types, handle weakened)
        //   • PythonDerivedType.Finalize    (Python-derived types, final cleanup)
        // ClassDerivedObject.ToPython re-adds entries when a deallocated wrapper
        // is resurrected with a strong GCHandle.
        internal static readonly HashSet<IntPtr> reflectedObjects = new();

        /// <summary>
        /// Current number of entries in the reflected-objects tracking set.
        /// Useful for diagnostics/monitoring memory behaviour.
        /// </summary>
        public static int ReflectedObjectCount => reflectedObjects.Count;

        /// <summary>
        /// Diagnostic: counts how many entries in reflectedObjects have each
        /// refcount value (1, 2, 3+).  Must hold GIL.
        /// </summary>
        internal static (int rc1, int rc2, int rc3plus) DiagnoseRefcounts()
        {
            int rc1 = 0, rc2 = 0, rc3plus = 0;
            foreach (var addr in reflectedObjects)
            {
                if (addr == IntPtr.Zero) continue;
                try
                {
                    var borrowed = new BorrowedReference(addr);
                    var rc = Runtime.Refcount(borrowed);
                    if (rc == 1) rc1++;
                    else if (rc == 2) rc2++;
                    else rc3plus++;
                }
                catch { }
            }
            return (rc1, rc2, rc3plus);
        }

        /// <summary>
        /// Diagnostic: returns a histogram of the .NET types currently in
        /// <see cref="reflectedObjects"/>.  Returns a list of
        /// <c>(typeFullName, count, totalRefcount, isPythonDerived,
        ///   parentlessSteps)</c> tuples sorted by count descending.
        /// <para>Must hold the Python GIL.</para>
        /// </summary>
        internal static IReadOnlyList<(string TypeName, int Count, long TotalRc, int PythonDerivedCount, int ParentlessStepCount)>
            DiagnoseTypeHistogram(int topN = 20)
        {
            var perType = new Dictionary<string, (int count, long rc, int pyDerived, int parentless)>(64);

            foreach (var addr in reflectedObjects)
            {
                if (addr == IntPtr.Zero) continue;
                try
                {
                    var borrowed = new BorrowedReference(addr);
                    var handle = ManagedType.TryGetGCHandle(borrowed);
                    if (handle is null || !handle.Value.IsAllocated) continue;
                    var target = handle.Value.Target;
                    if (target is not CLRObject clr) continue;

                    object inst = clr.inst;
                    if (inst is null) continue;

                    string typeName = inst.GetType().FullName ?? inst.GetType().Name;
                    long rc = Runtime.Refcount(borrowed);
                    int isDerived = inst is IPythonDerivedType ? 1 : 0;
                    int isParentless = 0;
                    // Classify orphaned ITestStep without taking a hard
                    // dependency on the OpenTAP types — duck-type via reflection
                    // on a "Parent" property.
                    try
                    {
                        var t = inst.GetType();
                        var parentProp = t.GetProperty("Parent");
                        if (parentProp != null)
                        {
                            var parent = parentProp.GetValue(inst);
                            if (parent is null) isParentless = 1;
                        }
                    }
                    catch { }

                    if (perType.TryGetValue(typeName, out var entry))
                        perType[typeName] = (entry.count + 1, entry.rc + rc, entry.pyDerived + isDerived, entry.parentless + isParentless);
                    else
                        perType[typeName] = (1, rc, isDerived, isParentless);
                }
                catch { }
            }

            return perType
                .OrderByDescending(kv => kv.Value.count)
                .Take(topN)
                .Select(kv => (kv.Key, kv.Value.count, kv.Value.rc, kv.Value.pyDerived, kv.Value.parentless))
                .ToList();
        }

        /// <summary>
        /// Scans the reflected-objects set for Python wrappers that are phantom
        /// references from <c>InvokeCtor</c> and can be safely released.
        /// <para>
        /// Three categories are handled:
        /// </para>
        /// <list type="number">
        ///   <item><b>Python-created phantoms</b> (<c>addr != __pyobj__</c>):
        ///     Orphaned wrappers from <c>InvokeCtor</c> that Python never uses.
        ///     Always safe to release (any refcount).</item>
        ///   <item><b>.NET-created owned wrappers</b> (<c>addr == __pyobj__</c>):
        ///     The primary wrapper for .NET-created PythonDerived instances.
        ///     Before releasing: <c>__dict__</c> is cleared (breaks cycles to
        ///     Trace wrappers etc.), then <c>__pyobj__</c> is zeroed, then the
        ///     phantom is released. Only evicted when <paramref name="isAbandoned"/>
        ///     confirms the object is no longer in use.</item>
        ///   <item><b>Stale CLR wrappers</b> (rc == 1, non-PythonDerived):
        ///     Regular CLR wrappers (e.g. TraceSource, LogEventType) that were
        ///     referenced from an evicted wrapper's <c>__dict__</c>. After the
        ///     dict-clearing cascade in category 2, these drop to rc = 1 and
        ///     can be safely released.</item>
        /// </list>
        /// <para><b>Must be called with the Python GIL held.</b></para>
        /// </summary>
        /// <param name="isAbandoned">
        /// Optional callback for .NET-created PythonDerived entries.
        /// Return <c>true</c> to allow eviction.
        /// If <c>null</c>, only Python-created phantoms are evicted.
        /// </param>
        /// <returns>Number of entries released.</returns>
        internal static int EvictAbandonedObjects(Func<object, bool>? isAbandoned = null)
        {
            // LEAK-DEMO: neutralized to reproduce the original (unpatched) leak.
            return 0;
            int count = reflectedObjects.Count;
            if (count == 0) return 0;

            IntPtr[] snapshot = new IntPtr[count];
            reflectedObjects.CopyTo(snapshot);

            int released = 0;

            // Record which entries are already at rc=1 BEFORE eviction.
            // Pass 2 must not touch these — they belong to still-alive objects.
            var preEvictRc1 = new HashSet<IntPtr>();
            foreach (var addr in snapshot)
            {
                if (addr == IntPtr.Zero) continue;
                try
                {
                    if (Runtime.Refcount(new BorrowedReference(addr)) == 1)
                        preEvictRc1.Add(addr);
                }
                catch { }
            }

            // ── Pass 1: Evict PythonDerived phantoms and owned wrappers ────
            // These may have rc > 1 due to __dict__ cross-references.
            // Clearing __dict__ first breaks cycles; then Py_DecRef releases
            // the phantom ref.  Cascade: tp_dealloc on freed wrappers reduces
            // rc on their dependents, leaving stale CLR wrappers at rc = 1
            // for Pass 2.
            foreach (var addr in snapshot)
            {
                if (addr == IntPtr.Zero) continue;
                if (!reflectedObjects.Contains(addr)) continue; // already freed by cascade

                try
                {
                    var borrowed = new BorrowedReference(addr);

                    GCHandle? handle = ManagedType.TryGetGCHandle(borrowed);
                    if (handle is null || !handle.Value.IsAllocated) continue;
                    if (handle.Value.Target is not CLRObject clrObj) continue;

                    object inst = clrObj.inst;
                    if (inst is not IPythonDerivedType derived) continue;

                    // Read __pyobj__ to classify the entry.
                    IntPtr pyObjAddr;
                    try { pyObjAddr = PythonDerivedType.GetPyObj(derived).RawObj; }
                    catch { continue; }

                    bool isPythonCreatedPhantom =
                        pyObjAddr != IntPtr.Zero && pyObjAddr != addr;

                    if (isPythonCreatedPhantom)
                    {
                        // Always safe — nobody uses this wrapper.
                        ClearDict(borrowed);
                        reflectedObjects.Remove(addr);
                        Runtime.XDecref(StolenReference.DangerousFromPointer(addr));
                        released++;
                    }
                    else if (isAbandoned != null && isAbandoned(inst))
                    {
                        // Owned wrapper for an abandoned .NET object.
                        // Clear __dict__ to cascade-free contained wrappers.
                        ClearDict(borrowed);
                        PythonDerivedType.SetPyObj(derived,
                            new BorrowedReference(IntPtr.Zero));
                        reflectedObjects.Remove(addr);
                        Runtime.XDecref(StolenReference.DangerousFromPointer(addr));
                        released++;
                    }
                }
                catch { /* skip corrupt entries */ }
            }

            // ── Pass 2: Sweep stale CLR wrappers now at rc = 1 ─────────
            // After Pass 1 cleared __dict__ on PythonDerived wrappers, their
            // contained CLR wrappers (TraceSource, enum values, etc.) may
            // have dropped to rc = 1.  Only release entries that were NOT
            // already rc = 1 before eviction (those are from still-alive objects).
            //
            // NOTE: Pass 2 always runs — even when Pass 1 found no abandoned
            // entries. Stale rc=1 wrappers can accumulate from earlier cleanup
            // cycles whose predicate did not match the current scenario; gating
            // Pass 2 on Pass 1's result would leak them indefinitely.
            {
                count = reflectedObjects.Count;
                snapshot = new IntPtr[count];
                reflectedObjects.CopyTo(snapshot);

                foreach (var addr in snapshot)
                {
                    if (addr == IntPtr.Zero) continue;
                    if (preEvictRc1.Contains(addr)) continue; // pre-existing, don't touch
                    try
                    {
                        var borrowed = new BorrowedReference(addr);
                        if (Runtime.Refcount(borrowed) != 1) continue;

                        reflectedObjects.Remove(addr);
                        Runtime.XDecref(StolenReference.DangerousFromPointer(addr));
                        released++;
                    }
                    catch { /* skip */ }
                }
            }

            return released;
        }

        /// <summary>
        /// Clears the Python <c>__dict__</c> on <paramref name="ob"/> to
        /// release any contained references before the wrapper is freed.
        /// </summary>
        private static void ClearDict(BorrowedReference ob)
        {
            try
            {
                using var dict = Runtime.PyObject_GenericGetDict(ob);
                if (!dict.IsNull())
                {
                    Runtime.PyDict_Clear(dict.Borrow());
                }
            }
            catch
            {
                // Dict may not exist for this type — ignore.
            }
        }

        static NewReference Create(object ob, BorrowedReference tp)
        {
            Debug.Assert(tp != null);
            var py = Runtime.PyType_GenericAlloc(tp, 0);

            var self = new CLRObject(ob);

            GCHandle gc = GCHandle.Alloc(self);
            InitGCHandle(py.Borrow(), type: tp, gc);

            bool isNew = reflectedObjects.Add(py.DangerousGetAddress());
            Debug.Assert(isNew);

            // Fix the BaseException args (and __cause__ in case of Python 3)
            // slot if wrapping a CLR exception
            if (ob is Exception e) Exceptions.SetArgsAndCause(py.Borrow(), e);

            return py;
        }

        CLRObject(object inst)
        {
            this.inst = inst;
        }

        /// <summary>
        /// Creates a new <see cref="CLRObject"/> wrapping <paramref name="inst"/>
        /// without allocating a new Python object or registering in
        /// <see cref="reflectedObjects"/>. Used by <c>ToPython</c> resurrection
        /// when the previous CLRObject was collected while its GCHandle was Weak.
        /// </summary>
        internal static CLRObject CreateWrapper(object inst) => new CLRObject(inst);

        internal static NewReference GetReference(object ob, BorrowedReference pyType)
            => Create(ob, pyType);

        internal static NewReference GetReference(object ob, Type type)
        {
            BorrowedReference cc = ClassManager.GetClass(type);
            return Create(ob, cc);
        }

        internal static NewReference GetReference(object ob)
        {
            BorrowedReference cc = ClassManager.GetClass(ob.GetType());
            return Create(ob, cc);
        }

        internal static void Restore(object ob, BorrowedReference pyHandle, Dictionary<string, object?> context)
        {
            var co = new CLRObject(ob);
            co.OnLoad(pyHandle, context);
        }

        protected override void OnLoad(BorrowedReference ob, Dictionary<string, object?>? context)
        {
            base.OnLoad(ob, context);
            GCHandle gc = GCHandle.Alloc(this);
            SetGCHandle(ob, gc);

            bool isNew = reflectedObjects.Add(ob.DangerousGetAddress());
            Debug.Assert(isNew);
        }

    }
}
