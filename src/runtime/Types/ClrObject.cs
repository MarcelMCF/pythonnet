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
        /// <summary>Number of entries with refcount ≤ 0 (zombies) that were evicted.</summary>
        public int EvictedZombies { get; init; }
        /// <summary>Number of entries that caused an access violation and were evicted.</summary>
        public int EvictedInvalid { get; init; }
        /// <summary>Number of entries that are still alive (refcount > 1).</summary>
        public int Alive { get; init; }
        /// <summary>Total entries in reflectedObjects before eviction.</summary>
        public int TotalBefore { get; init; }
        /// <summary>Total entries in reflectedObjects after eviction.</summary>
        public int TotalAfter { get; init; }
        /// <summary>Total number of evicted entries.</summary>
        public int TotalEvicted => EvictedRc1 + EvictedZombies + EvictedInvalid;
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
        /// Entries with Python refcount = 1 are only referenced by this tracking set.
        /// Their GCHandle is freed (releasing the pinned .NET object), and they are removed
        /// from the set. Entries with refcount ≤ 0 (zombies from previous frees) are also
        /// removed. Entries with refcount > 1 are still in use by Python code and are left
        /// untouched.
        /// </para>
        /// <para>
        /// <b>Important:</b> The caller must hold the GIL. No <c>Py_DecRef</c> is called —
        /// the Python-side objects remain alive but the .NET GCHandle is freed, allowing the
        /// .NET garbage collector to reclaim the managed object. Python.Runtime will create
        /// fresh CLRObject wrappers if needed later.
        /// </para>
        /// </summary>
        /// <returns>An <see cref="EvictResult"/> with eviction statistics.</returns>
        public static EvictResult EvictCollectable()
        {
            int evictedRc1 = 0, evictedZombies = 0, evictedInvalid = 0, alive = 0;
            int totalBefore = reflectedObjects.Count;

            var snapshot = new IntPtr[reflectedObjects.Count];
            reflectedObjects.CopyTo(snapshot);

            var toEvict = new List<IntPtr>();

            // Pass 1: classify by refcount (read-only)
            foreach (var ptr in snapshot)
            {
                if (ptr == IntPtr.Zero) { toEvict.Add(ptr); evictedInvalid++; continue; }

                try
                {
                    var borrowed = new BorrowedReference(ptr);
                    nint refcount = Runtime.Refcount(borrowed);

                    if (refcount <= 0)
                    {
                        evictedZombies++;
                        toEvict.Add(ptr);
                    }
                    else if (refcount == 1)
                    {
                        evictedRc1++;
                        toEvict.Add(ptr);
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
                    toEvict.Add(ptr);
                }
            }

            // Pass 2: free GCHandles and remove from tracking set
            foreach (var ptr in toEvict)
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

            return new EvictResult
            {
                EvictedRc1 = evictedRc1,
                EvictedZombies = evictedZombies,
                EvictedInvalid = evictedInvalid,
                Alive = alive,
                TotalBefore = totalBefore,
                TotalAfter = reflectedObjects.Count,
            };
        }
    }
}
