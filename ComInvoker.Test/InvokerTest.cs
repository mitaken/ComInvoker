using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ComInvoker.Test
{
    [TestClass]
    public class InvokerTest
    {
        private const string TestInProcess = nameof(TestInProcess);
        private const string TestOutProcess = nameof(TestOutProcess);
        const int READYSTATE_COMPLETE = 4;

        [TestMethod]
        [TestCategory(TestInProcess)]
        public void TestInProcessAutoRelase()
        {
            VBScript_RegExp_55.RegExp regex;
            using (var invoker = new Invoker())
            {
                regex = invoker.Invoke<VBScript_RegExp_55.RegExp>(new VBScript_RegExp_55.RegExp());
                regex.Pattern = "^[0-9]$";
                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(regex.Test("1"));

                IEnumerable<VBScript_RegExp_55.Match> matches = invoker.InvokeEnumurator<VBScript_RegExp_55.Match>(regex.Execute("2"));
                foreach (var match in matches)
                {
                    Assert.AreEqual("2", match.Value);
                }
                Assert.AreEqual(3, invoker.StackCount);

            }
            Assert.ThrowsException<InvalidComObjectException>(() => regex.Test("1"), "Release faiure");
        }


        [TestMethod]
        [TestCategory(TestInProcess)]
        public void TestInProcessManualRelease()
        {
            using (var invoker = new Invoker())
            {
                var regex = invoker.Invoke<VBScript_RegExp_55.RegExp>(new VBScript_RegExp_55.RegExp());
                regex.Pattern = "^[0-9]$";
                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(regex.Test("1"));

                IEnumerable<VBScript_RegExp_55.Match> matches = invoker.InvokeEnumurator<VBScript_RegExp_55.Match>(regex.Execute("2"));
                foreach (var match in matches)
                {
                    Assert.AreEqual("2", match.Value);
                }
                Assert.AreEqual(3, invoker.StackCount);

                invoker.Release(3);

                Assert.AreEqual(0, invoker.StackCount);
                Assert.ThrowsException<InvalidComObjectException>(() => regex.Pattern, "Release faiure");
            }
        }

        [TestMethod]
        [TestCategory(TestOutProcess)]
        public void TestOutProcessAutoRelease()
        {
            var type = Type.GetTypeFromProgID("InternetExplorer.Application");

            dynamic ie;
            using (var invoker = new Invoker())
            {
                ie = invoker.Invoke<dynamic>(Activator.CreateInstance(type));
                ie.Visible = true;

                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(ie.Visible);

                ie.Navigate("about:blank");

                while (ie.Busy || ie.ReadyState != READYSTATE_COMPLETE) { }
                ie.Quit();
            }
            Assert.ThrowsException<InvalidComObjectException>(() => ie.Visible, "Release faiure");
        }

        [TestMethod]
        [TestCategory(TestOutProcess)]
        public void TestOutProcessManualRelease()
        {
            var type = Type.GetTypeFromProgID("InternetExplorer.Application");

            using (var invoker = new Invoker())
            {
                var ie = invoker.Invoke<dynamic>(Activator.CreateInstance(type));
                ie.Visible = true;

                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(ie.Visible);

                ie.Navigate("about:blank");

                while (ie.Busy || ie.ReadyState != READYSTATE_COMPLETE) { }
                ie.Quit();

                invoker.Release();

                Assert.AreEqual(0, invoker.StackCount);
                Assert.ThrowsException<InvalidComObjectException>(() => ie.Pattern, "Release faiure");
            }
        }
    }
}
