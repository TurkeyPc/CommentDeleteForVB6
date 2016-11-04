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
            var actual = Program.DeleteComment("s = \"---I'm a student.---\"");
            var expected = "s = \"---I'm a student.---\"";

            Assert.AreEqual(expected, actual);
        }
    }
}
