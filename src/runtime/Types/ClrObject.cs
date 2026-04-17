using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Python.Runtime
{
    /// <summary>
    /// Result of a <see cref="CLRObject.EvictCollectable"/> call.
    /// </summary>
    public readonly struct EvictResult
    {
        /// <summary>Number of entries with refcount ≤ 1 that were evicted and had their GCHandle freed.</summary>
        public int EvictedRc1 { get; init; }
        /// <summary>Number of entries with refcount &gt; 1 that were force-evicted (GCHandle freed, Python wrapper kept alive).</summary>
        public int EvictedAlive { get; init; }
        /// <summary>Number of entries with refcount ≤ 0 (zombies) that were evicted.</summary>
        public int EvictedZombies { get; init; }
        /// <summary>Number of entries that caused an access violation and were evicted.</summary>
        public int EvictedInvalid { get; init; }
        /// <summary>Number of entries that are still alive (refcount &gt; maxRefcount) and were left untouched.</summary>
        public int Alive { get; init; }
        /// <summary>Total entries in reflectedObjects before eviction.</summary>
        public int TotalBefore { get; init; }
        /// <summary>Total entries in reflectedObjects after eviction.</summary>
        public int TotalAfter { get; init; }
        /// <summary>Total number of evicted entries.</summary>
        public int TotalEvicted => EvictedRc1 + EvictedAlive + EvictedZombies + EvictedInvalid;
    }

    [Serializable]
    [DebuggerDisplay("clrO: {inst}")]
    internal sealed class CLRObject : ManagedType
    {
        internal readonly object inst;

        // "borrowed" references
        internal static readonly HashSet<IntPtr> reflectedObjects = new();

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

        /// <summary>
        /// Evicts collectable entries from <see cref="reflectedObjects"/>.
        /// <para>
        /// Entries with Python refcount ≤ 1 are only referenced by this tracking set.
        /// Their GCHandle is freed (releasing the pinned .NET object), and they are removed
        /// from the set. Entries with refcount ≤ 0 (zombies from previous frees) are also
        /// removed.
        /// </para>
        /// <para>
        /// When <paramref name="maxRefcount"/> is greater than 1, entries with refcount &gt; 1
        /// are handled based on their managed object type:
        /// <list type="bullet">
        ///   <item><b>Python-derived instances</b> (<see cref="IPythonDerivedType"/>): A single
        ///   <c>Py_DecRef</c> is called to release the phantom reference leaked by
        ///   <c>ClassDerived.InvokeCtor</c>. Since Entry refcount was &gt; 1, after DecRef
        ///   it stays ≥ 1 — the Python object survives, <c>tp_dealloc</c> does not fire,
        ///   and the GCHandle remains Strong. The <c>__pyobj__</c> field is intentionally
        ///   left intact so the .NET step can still access its Python backing object
        ///   (required when a test plan is reused across multiple runs, e.g. Session 2).
        ///   Normal finalization (<c>PyFinalize</c> → <c>PyObject_GC_Del</c>) will clean up
        ///   when the .NET step is eventually GC'd.</item>
        ///   <item><b>Infrastructure objects</b> (ExtensionType, MethodObject, type wrappers):
        ///   Removed from the tracking set <b>without touching their GCHandle or refcount</b>.
        ///   These are reused by Python's type system across runs — touching them causes access
        ///   violations.</item>
        /// </list>
        /// </para>
        /// <para>
        /// <b>Important:</b> The caller must hold the GIL.
        /// </para>
        /// </summary>
        /// <param name="maxRefcount">
        /// Maximum Python refcount threshold for eviction. Objects with refcount ≤ 1 have
        /// their GCHandle freed. Objects with 1 &lt; refcount ≤ maxRefcount are processed
        /// based on type (Py_DecRef for PythonDerived, untrack-only for infrastructure).
        /// Default is 1 (conservative: only evicts objects not in active use).
        /// </param>
        /// <returns>An <see cref="EvictResult"/> with eviction statistics.</returns>
        public static EvictResult EvictCollectable(long maxRefcount = 1)
        {
            int evictedRc1 = 0, evictedAlive = 0, evictedZombies = 0, evictedInvalid = 0, alive = 0;
            int totalBefore = reflectedObjects.Count;

            // Trigger Python's cyclic GC to break reference cycles and potentially
            // reduce refcounts before we classify objects.
            try { Runtime.PyGC_Collect(); } catch { /* best-effort */ }

            var snapshot = new IntPtr[reflectedObjects.Count];
            reflectedObjects.CopyTo(snapshot);

            var toFreeGCHandle = new List<IntPtr>();   // rc ≤ 1: free GCHandle directly
            var toDecref = new List<IntPtr>();          // rc > 1 + PythonDerived: Py_DecRef phantom ref
            var toUntrack = new List<IntPtr>();         // rc > 1 + infrastructure: remove from set only

            // Pass 1: classify by refcount and managed object type (read-only)
            foreach (var ptr in snapshot)
            {
                if (ptr == IntPtr.Zero) { toFreeGCHandle.Add(ptr); evictedInvalid++; continue; }

                try
                {
                    var borrowed = new BorrowedReference(ptr);
                    nint refcount = Runtime.Refcount(borrowed);

                    if (refcount <= 0)
                    {
                        evictedZombies++;
                        toFreeGCHandle.Add(ptr);
                    }
                    else if (refcount == 1)
                    {
                        evictedRc1++;
                        toFreeGCHandle.Add(ptr);
                    }
                    else if (refcount <= maxRefcount)
                    {
                        // rc > 1: check if this is a PythonDerived instance (step) or
                        // infrastructure (type wrapper, method descriptor, etc.).
                        bool isPythonDerived = false;
                        try
                        {
                            var gcHandle = TryGetGCHandle(borrowed);
                            if (gcHandle is { IsAllocated: true })
                            {
                                var target = gcHandle.Value.Target;
                                if (target is CLRObject clrObj && clrObj.inst is IPythonDerivedType)
                                {
                                    isPythonDerived = true;
                                }
                            }
                        }
                        catch
                        {
                            // GCHandle access failed — treat as infrastructure (safe path)
                        }

                        if (isPythonDerived)
                        {
                            // PythonDerived step instance: release the phantom reference
                            // from InvokeCtor. ClassDerived.tp_dealloc (if rc→0) safely
                            // downgrades the GCHandle to Weak without freeing Python memory.
                            evictedAlive++;
                            toDecref.Add(ptr);
                        }
                        else
                        {
                            // Infrastructure object: just untrack. Touching GCHandle or
                            // refcount of ExtensionType/MethodObject/etc. causes crashes.
                            evictedAlive++;
                            toUntrack.Add(ptr);
                        }
                    }
                    else
                    {
                        alive++;
                    }
                }
                catch
                {
                    // Pointer is no longer valid
                    evictedInvalid++;
                    toFreeGCHandle.Add(ptr);
                }
            }

            // Pass 2a: free GCHandles for rc ≤ 1 objects and remove from tracking set
            foreach (var ptr in toFreeGCHandle)
            {
                if (ptr != IntPtr.Zero)
                {
                    try
                    {
                        var borrowed = new BorrowedReference(ptr);
                        TryFreeGCHandle(borrowed);
                    }
                    catch
                    {
                        // Best-effort — pointer may be stale
                    }
                }
                reflectedObjects.Remove(ptr);
            }

            // Pass 2b: Py_DecRef for PythonDerived step instances.
            // This releases the phantom reference leaked by InvokeCtor:
            //   var pyRef = CLRObject.GetReference(obj, type);  // rc = 1
            //   SetPyObj(obj, pyRef.Borrow());
            //   // pyRef never disposed — phantom ref leaked
            //
            // Since these entries had refcount > 1 at classification time, after
            // Py_DecRef the refcount stays ≥ 1 — the Python object survives and
            // tp_dealloc does NOT fire. The GCHandle remains Strong, so .NET GC
            // cannot collect the step prematurely.
            //
            // We intentionally do NOT null __pyobj__: the .NET step's Python backing
            // object is still alive (rc ≥ 1), and nulling would break any subsequent
            // property access — e.g. when Session 2 reuses the same test plan across
            // multiple runs. Normal finalization (PyFinalize → PyObject_GC_Del) will
            // clean up when the step is eventually GC'd.
            foreach (var ptr in toDecref)
            {
                try
                {
                    reflectedObjects.Remove(ptr);
                    Runtime.Py_DecRef(StolenReference.DangerousFromPointer(ptr));
                }
                catch
                {
                    // Best-effort — Py_DecRef may trigger complex chains
                    reflectedObjects.Remove(ptr);
                }
            }

            // Pass 2c: remove infrastructure objects from tracking set only.
            // GCHandle and refcount are left untouched — these objects are managed
            // by Python's type system and will be cleaned up at shutdown.
            foreach (var ptr in toUntrack)
            {
                reflectedObjects.Remove(ptr);
            }

            return new EvictResult
            {
                EvictedRc1 = evictedRc1,
                EvictedAlive = evictedAlive,
                EvictedZombies = evictedZombies,
                EvictedInvalid = evictedInvalid,
                Alive = alive,
                TotalBefore = totalBefore,
                TotalAfter = reflectedObjects.Count,
            };
        }

    }
}
