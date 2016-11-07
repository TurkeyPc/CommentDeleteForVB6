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
            var target = new VB6Source(new string[] { "Dim v As Integer 'v is Integer" });
            var actual = target.CommentDeleted.ToArray();
            var expected = new[] { "Dim v As Integer " };

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestStartSideComment()
        {
            var target = new VB6Source(new string[] { "'    'Dim v As Integer 'v is Integer"});
            var actual = target.CommentDeleted.ToArray();
            var expected = new string[0];

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestCommentMarkInString()
        {
            var s1 = Properties.Resources.TestMethod3InputValue;
            /* Value of s1 is
            s = "---I'm a student.---"
            */

            var target = new VB6Source(new string[] { s1 });
            var actual = target.CommentDeleted.ToArray();
            var expected = new[] { s1 };

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestCommentMarkInStringWithDoubleQuote()
        {
            var s1 = Properties.Resources.TestMethod4InputValue;
            /* Value of s1 is
            s = "---""I'm a student.""---"
            */

            var target = new VB6Source(new string[] { s1 });
            var actual = target.CommentDeleted.ToArray();
            var expected = new[] { s1 };

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithoutComment()
        {
            var s1 = Properties.Resources.TestMehtod5InputValue;
            /* Value of s1 is
            Dim s As _
            String
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            var actual = target.CommentDeleted.ToArray();
            var expected = s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithEndComment()
        {
            var s1 = Properties.Resources.TestMehtod6InputValue;
            /* Value of s1 is
            Dim s As _
            String 's Is String
            */

            var s2 = Properties.Resources.TestMehtod6ExpectedValue;
            /* Value of s1 is
            Dim s As _
            String 
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            var actual = target.CommentDeleted.ToArray();
            var expected = s2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithStartComment()
        {
            var s1 = Properties.Resources.TestMehtod7InputValue;
            /* Value of s1 is
            'Dim s As _
            String
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            var actual = target.CommentDeleted.ToArray();
            var expected = new string[0];

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestMultiLineWithCenterComment()
        {
            var s1 = Properties.Resources.TestMehtod8InputValue;
            /* Value of s1 is
            Dim s As _
            String 's Is String _
            s = "New Value"
            */

            var s2 = Properties.Resources.TestMehtod6ExpectedValue; //expected is same as TestMultiLineWithEndComment
            /* Value of s1 is
            Dim s As _
            String 
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            var actual = target.CommentDeleted.ToArray();
            var expected = s2.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestRealVB6Method1()
        {
            var s1 = Properties.Resources.TestMethod9InputValue;
            /* Value of s1 is
            Private Sub Form_Load()

            'Dim s As Integer _
            Dim v As Variant

            End Sub
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
            var actual = target.CommentDeleted.ToArray();
            var expected = new List<string>();
            expected.Add("Private Sub Form_Load()");
            expected.Add("End Sub");

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void TestRealVB6Method2()
        {
            var s1 = Properties.Resources.TestMethodAInputValue;
            /* Value of s1 is
            Private Sub Form_Load()

            'Dim s As Integer _

            Dim v As Variant

            End Sub
            */

            var target = new VB6Source(s1.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
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
