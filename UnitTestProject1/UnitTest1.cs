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

        [TestMethod]
        public void TestMethod9()
        {
            var s = Properties.Resources.TestMethod9InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = (Program.LogicalRows(ss)).ToArray();

            Assert.AreEqual(3, actual.Count());
            Assert.AreEqual(1, actual[0].Count());
            Assert.AreEqual(2, actual[1].Count());
            Assert.AreEqual(1, actual[2].Count());

            Assert.AreEqual("Private Sub Form_Load()", actual[0][0]);
            Assert.AreEqual("'Dim s As Integer _", actual[1][0]);
            Assert.AreEqual("Dim v As Variant", actual[1][1]);
            Assert.AreEqual("End Sub", actual[2][0]);
        }


        [TestMethod]
        public void TestMethodA()
        {
            var s = Properties.Resources.TestMethodAInputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = (Program.LogicalRows(ss)).ToArray();

            Assert.AreEqual(4, actual.Count());
            Assert.AreEqual(1, actual[0].Count());
            Assert.AreEqual(1, actual[1].Count());
            Assert.AreEqual(1, actual[2].Count());
            Assert.AreEqual(1, actual[3].Count());

            Assert.AreEqual("Private Sub Form_Load()", actual[0][0]);
            Assert.AreEqual("'Dim s As Integer _", actual[1][0]);

            Assert.AreEqual("Dim v As Variant", actual[2][0]);
            Assert.AreEqual("End Sub", actual[3][0]);
        }

        [TestMethod]
        public void TestMethodCombo1()
        {
            var s = Properties.Resources.TestMethod9InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = Program.LogicalRows(ss).Select(p => Program.DeleteComment2(p)).ToArray();

            Assert.AreEqual(3, actual.Count());
            Assert.AreEqual(1, actual[0].Count());
            Assert.AreEqual(1, actual[1].Count());
            Assert.AreEqual(1, actual[2].Count());

            Assert.AreEqual("Private Sub Form_Load()", actual[0].ToArray()[0]);
            Assert.AreEqual("", actual[1].ToArray()[0]);
            Assert.AreEqual("End Sub", actual[2].ToArray()[0]);
        }

        [TestMethod]
        public void TestMethodCombo2()
        {
            var s = Properties.Resources.TestMethodAInputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = Program.LogicalRows(ss).Select(p => Program.DeleteComment2(p)).ToArray();

            Assert.AreEqual(4, actual.Count());
            Assert.AreEqual(1, actual[0].Count());
            Assert.AreEqual(1, actual[1].Count());
            Assert.AreEqual(1, actual[2].Count());
            Assert.AreEqual(1, actual[3].Count());

            Assert.AreEqual("Private Sub Form_Load()", actual[0].ToArray()[0]);
            Assert.AreEqual("", actual[1].ToArray()[0]);
            Assert.AreEqual("Dim v As Variant", actual[2].ToArray()[0]);
            Assert.AreEqual("End Sub", actual[3].ToArray()[0]);
        }


        [TestMethod]
        public void TestMethodCombo3()
        {
            var s = Properties.Resources.TestMehtod8InputValue;
            var ss = s.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var actual = Program.LogicalRows(ss).Select(p => Program.DeleteComment2(p)).ToArray();

            Assert.AreEqual(1, actual.Count());
            Assert.AreEqual(2, actual[0].Count());

            Assert.AreEqual("Dim s As _", actual[0].ToArray()[0]);
            Assert.AreEqual("String ", actual[0].ToArray()[1]);
        }
    }
}
