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
        /// Invoke COM object
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
        /// Invoke Enumurator
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
        /// Release COM object
        /// </summary>
        /// <param name="releaseCount">Release count</param>
        public bool[] Release(int releaseCount = 1)
        {
            var results = new bool[releaseCount];
            for (var i = 0; i < releaseCount; i++)
            {
                var result = false;
                if (comStack.TryPop(out object com))
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
