using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ComInvoker
{
    /// <summary>
    /// COM Invoker
    /// </summary>
    public class Invoker : IDisposable
    {
        /// <summary>
        /// COM stacker
        /// </summary>
        private ConcurrentStack<object> comStack = new ConcurrentStack<object>();

        /// <summary>
        /// COM stacked count
        /// </summary>
        public int StackCount
        {
            get { return comStack.Count; }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        public Invoker()
        { }

        /// <summary>
        /// Invoke COM object from ProgID (early binding)
        /// </summary>
        /// <typeparam name="T">Type of COM</typeparam>
        /// <param name="progID">Programmatic Identifier</param>
        /// <returns>Typed COM</returns>
        public T InvokeFromProgID<T>(string progID)
        {
            return Invoke<T>(Activator.CreateInstance(Type.GetTypeFromProgID(progID, true)));
        }

        /// <summary>
        /// Invoke COM object from ProgID (late binding)
        /// </summary>
        /// <param name="progID">Programmatic Identifier</param>
        /// <returns>Dynamic COM</returns>
        public dynamic InvokeFromProgID(string progID)
        {
            return InvokeFromProgID<dynamic>(progID);
        }

        /// <summary>
        /// Invoke COM object (early binding)
        /// </summary>
        /// <remarks>Thread safe</remarks>
        /// <typeparam name="T">Type of COM</typeparam>
        /// <param name="com">COM object</param>
        /// <returns>Typed COM</returns>
        public T Invoke<T>(object com)
        {
            comStack.Push(com);

            return (T)com;
        }

        /// <summary>
        /// Invoke COM object (late binding)
        /// </summary>
        /// <param name="com">Type of COM</param>
        /// <returns>Dynamic COM</returns>
        public dynamic Invoke(object com)
        {
            return Invoke<dynamic>(com);
        }

        /// <summary>
        /// Invoke Enumurator (early binding)
        /// </summary>
        /// <typeparam name="T">Type of COM</typeparam>
        /// <param name="com">COM object</param>
        /// <returns>Typed COM</returns>
        public IEnumerable<T> InvokeEnumurator<T>(IEnumerable com)
        {
            return Invoke<IEnumerable>(com)
                .Cast<T>()
                .Select(x => Invoke<T>(x));
        }

        /// <summary>
        /// Invoke Enumurator (late binding)
        /// </summary>
        /// <param name="com">COM object</param>
        /// <returns>Dynamic COM</returns>
        public IEnumerable<dynamic> InvokeEnumurator(IEnumerable com)
        {
            return InvokeEnumurator<dynamic>(com);
        }

        /// <summary>
        /// Release COM object
        /// </summary>
        /// <param name="releaseCount">Release count</param>
        public bool[] Release(int releaseCount = 1)
        {
            var results = new bool[releaseCount];
            for (var i = 0; i < releaseCount; i++)
            {
                var result = false;
                if (comStack.TryPop(out var com))
                {
                    result = InternalRelease(com);
                }
                results[i] = result;
            }

            return results;
        }

        /// <summary>
        /// Release all COM objects
        /// </summary>
        public bool[] ReleaseAll()
        {
            return Release(comStack.Count);
        }

        /// <summary>
        /// Internal release method
        /// </summary>
        /// <param name="com">COM object</param>
        /// <returns>Release result</returns>
        private bool InternalRelease(object com)
        {
            var result = false;
            if (com != null && Marshal.IsComObject(com))
            {
                result = Marshal.FinalReleaseComObject(com) == 0;
                com = null;
            }

            return result;
        }

        #region IDisposable Support
        private bool disposedValue = false;

        /// <summary>
        /// Dispose com objects
        /// </summary>
        /// <param name="disposing">true:managed resource, false:unmanaged resource</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    ReleaseAll();
                }

                disposedValue = true;
            }
        }

        void IDisposable.Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
