using System;
using System.Collections;
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
        private Stack<object> comStack = new Stack<object>();

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
        public void Release(int releaseCount = 1)
        {
            for (int i = 0; i < releaseCount; i++)
            {
                InternalRelease(comStack.Pop());
            }
        }

        /// <summary>
        /// Release all COM objects
        /// </summary>
        public void ReleaseAll()
        {
            while (comStack.Count > 0)
            {
                InternalRelease(comStack.Pop());
            }
        }

        /// <summary>
        /// Internal release method
        /// </summary>
        /// <param name="com">COM object</param>
        private void InternalRelease(object com)
        {
            if (com != null && Marshal.IsComObject(com))
            {
                while (Marshal.FinalReleaseComObject(com) != 0) { }
                com = null;
            }
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
