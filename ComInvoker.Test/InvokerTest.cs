using Microsoft.VisualBasic;
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
        private const string TestInvoke = nameof(TestInvoke);
        private const string RegExpTypeName = "IRegExp2";
        private const string RegExpProgID = "VBScript.RegExp";
        private const string InternetExplorerProgID = "InternetExplorer.Application";


        const int READYSTATE_COMPLETE = 4;

        [TestMethod]
        [TestCategory(TestInvoke)]
        public void TestInvokeObject()
        {
            using (var invoker = new Invoker())
            {
                Assert.AreEqual(Information.TypeName(invoker.Invoke<VBScript_RegExp_55.RegExp>(new VBScript_RegExp_55.RegExp())), RegExpTypeName);
                Assert.AreEqual(Information.TypeName(invoker.Invoke(new VBScript_RegExp_55.RegExp())), RegExpTypeName);
            }
        }

        [TestMethod]
        [TestCategory(TestInvoke)]
        public void TestInvokeFromProgID()
        {
            using (var invoker = new Invoker())
            {
                Assert.AreEqual(Information.TypeName(invoker.InvokeFromProgID<VBScript_RegExp_55.RegExp>(RegExpProgID)), RegExpTypeName);
                Assert.AreEqual(Information.TypeName(invoker.InvokeFromProgID(RegExpProgID)), RegExpTypeName);

                Assert.ThrowsException<COMException>(() => invoker.InvokeFromProgID<Action>("dummy_early_binding_com_objects"));
                Assert.ThrowsException<COMException>(() => invoker.InvokeFromProgID("dummy_late_binding_com_objects"));
            }
        }

        [TestMethod]
        [TestCategory(TestInProcess)]
        public void TestInProcessAutoRelase()
        {
            VBScript_RegExp_55.RegExp regex;
            using (var invoker = new Invoker())
            {
                regex = invoker.InvokeFromProgID<VBScript_RegExp_55.RegExp>(RegExpProgID);
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
                var regex = invoker.InvokeFromProgID<VBScript_RegExp_55.RegExp>(RegExpProgID);
                regex.Pattern = "^[0-9]$";
                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(regex.Test("1"));

                IEnumerable<VBScript_RegExp_55.Match> matches = invoker.InvokeEnumurator<VBScript_RegExp_55.Match>(regex.Execute("2"));
                foreach (var match in matches)
                {
                    Assert.AreEqual("2", match.Value);
                }
                Assert.AreEqual(3, invoker.StackCount);

                CollectionAssert.DoesNotContain(invoker.ReleaseAll(), false);

                Assert.AreEqual(0, invoker.StackCount);
                Assert.ThrowsException<InvalidComObjectException>(() => regex.Pattern, "Release faiure");
            }
        }

        [TestMethod]
        [TestCategory(TestOutProcess)]
        public void TestOutProcessAutoRelease()
        {
            dynamic ie;
            using (var invoker = new Invoker())
            {
                ie = invoker.InvokeFromProgID(InternetExplorerProgID);
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
            using (var invoker = new Invoker())
            {
                var ie = invoker.InvokeFromProgID(InternetExplorerProgID);
                ie.Visible = true;

                Assert.AreEqual(1, invoker.StackCount);
                Assert.IsTrue(ie.Visible);

                ie.Navigate("about:blank");

                while (ie.Busy || ie.ReadyState != READYSTATE_COMPLETE) { }
                ie.Quit();

                CollectionAssert.DoesNotContain(invoker.Release(), false);

                Assert.AreEqual(0, invoker.StackCount);
                Assert.ThrowsException<InvalidComObjectException>(() => ie.Pattern, "Release faiure");
            }
        }
    }
}
