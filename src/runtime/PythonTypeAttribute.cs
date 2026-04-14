using System;

namespace Python.Runtime
{
    /// <summary>
    /// Applied to IL-generated properties on Python-derived CLR types when the
    /// property type resolves to <see cref="PyObject"/> (i.e. the Python type has
    /// no direct CLR equivalent). Stores the originating Python module and type
    /// name so higher-level code can resolve the Python type at runtime.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class PythonTypeAttribute : Attribute
    {
        /// <summary>Python type name (e.g. "MyEnum").</summary>
        public string TypeName { get; }

        /// <summary>Python module containing the type (e.g. "my_plugin").</summary>
        public string Module { get; }

        public PythonTypeAttribute(string pyTypeModule, string pyTypeName)
        {
            TypeName = pyTypeName;
            Module = pyTypeModule;
        }
    }
}
