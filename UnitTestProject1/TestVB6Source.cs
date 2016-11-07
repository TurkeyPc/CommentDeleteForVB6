using Microsoft.VisualStudio.TestTools.UnitTesting;
using CommentDeleteForVB6;
using System.Linq;
using System;
using System.Collections.Generic;

namespace UnitTestProject1
{
    [TestClass]
    public class TestVB6Source
    {
        [TestMethod]
        public void TestEndSideComment()
        {
            var s1 = new List<string>();
            s1.Add("Dim v As Integer 'v is Integer");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("Dim v As Integer ");

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestStartSideComment()
        {
            var s1 = new List<string>();
            s1.Add("'    'Dim v As Integer 'v is Integer");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new string[0];

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestCommentMarkInString()
        {
            var s1 = spliter(Properties.Resources.TestCommentMarkInStringInputValue);
            /* Value of s1 is
            s = "---I'm a student.---"
            */

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>(s1);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestCommentMarkInStringWithDoubleQuote()
        {
            var s1 = spliter(Properties.Resources.TestCommentMarkInStringWithDoubleQuoteInputValue);
            /* Value of s1 is
            s = "---""I'm a student.""---"
            */

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>(s1);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithoutComment()
        {
            var s1 = new List<string>();
            s1.Add("Dim s As _");
            s1.Add("String");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>(s1);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithEndComment()
        {
            var s1 = new List<string>();
            s1.Add("Dim s As _");
            s1.Add("String 's Is String");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("Dim s As _");
            expected.Add("String ");

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithStartComment()
        {
            var s1 = new List<string>();
            s1.Add("'Dim s As _");
            s1.Add("String");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();

            CollectionAssert.AreEqual(expected, actual);
        }

        private string[] spliter(string p)
        {
            return p.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        }

        [TestMethod]
        public void TestMultiLineWithCenterComment()
        {
            var s1 = spliter(Properties.Resources.TestMultiLineWithCenterCommentInputValue);
            /* Value of s1 is
            Dim s As _
            String 's Is String _
            s = "New Value"
            */

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("Dim s As _");
            expected.Add("String ");

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestRealVB6Method1()
        {
            var s1 = new List<string>();
            s1.Add("Private Sub Form_Load()");
            s1.Add("");
            s1.Add("'Dim s As Integer _");
            s1.Add("Dim v As Variant");
            s1.Add("");
            s1.Add("End Sub");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("Private Sub Form_Load()");
            expected.Add("End Sub");

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestRealVB6Method2()
        {
            var s1 = new List<string>();
            s1.Add("Private Sub Form_Load()");
            s1.Add("");
            s1.Add("'Dim s As Integer _");
            s1.Add("");
            s1.Add("Dim v As Variant");
            s1.Add("");
            s1.Add("End Sub");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("Private Sub Form_Load()");
            expected.Add("Dim v As Variant");
            expected.Add("End Sub");

            CollectionAssert.AreEqual(expected, actual);
        }
        

        [TestMethod]
        public void TestRealVB6Method3()
        {
            var s1 = new List<string>();
            s1.Add("For i = 0 To 10");
            s1.Add("    Debug.Print CStr(i) 'OK_");
            s1.Add("Next i");

            var target = new VB6Source(s1);
            var actual = target.CommentDeleted.ToArray();

            var expected = new List<string>();
            expected.Add("For i = 0 To 10");
            expected.Add("    Debug.Print CStr(i) ");
            expected.Add("Next i");

            CollectionAssert.AreEqual(expected, actual);
        }
    }
}
