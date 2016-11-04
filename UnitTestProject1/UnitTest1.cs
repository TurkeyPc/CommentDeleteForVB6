using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CommentDeleteForVB6;
using System.Linq;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var actual = Program.DeleteComment("Dim v As Integer 'v is Integer");
            var expected = "Dim v As Integer ";

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod2()
        {
            var actual = Program.DeleteComment("'    'Dim v As Integer 'v is Integer");
            var expected = "";

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod3()
        {
            var s = Properties.Resources.TestMethod3InputValue;

            var actual = Program.DeleteComment(s);
            var expected = Properties.Resources.TestMethod3ExpectedValue;

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod4()
        {
            var s = Properties.Resources.TestMethod4InputValue;

            var actual = Program.DeleteComment(s);
            var expected = Properties.Resources.TestMethod4ExpectedValue;

            Assert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod5()
        {
            var s = Properties.Resources.TestMehtod5InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var expected = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var actual = (Program.DeleteComment2(ss)).ToArray();

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod6()
        {
            var s = Properties.Resources.TestMehtod6InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = (Program.DeleteComment2(ss)).ToArray();
            var expected =Properties.Resources.TestMehtod6ExpectedValue.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod7()
        {
            var s = Properties.Resources.TestMehtod7InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = (Program.DeleteComment2(ss)).ToArray();
            var expected = new[] { "" };

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMethod8()
        {
            var s = Properties.Resources.TestMehtod8InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = (Program.DeleteComment2(ss)).ToArray();
            var expected =Properties.Resources.TestMehtod6ExpectedValue.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            CollectionAssert.AreEqual(expected, actual);
        }

    }
}
