using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CommentDeleteForVB6;

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
    }
}
