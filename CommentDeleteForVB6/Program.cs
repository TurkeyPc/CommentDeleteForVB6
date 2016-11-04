using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CommentDeleteForVB6
{
    public static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        public static string DeleteComment(string s)
        {
            var c = s.IndexOf("'");

            if (c == -1)
                return s;

            if (s.IndexOf("\"") == -1)
                return s.Substring(0, c);

            bool InString = false;

            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '"')
                    InString = !InString;

                if (!InString && s[i] == '\'')
                    return s.Substring(0, i - 1);
            }

            return s;
        }

        public static IEnumerable<string> DeleteComment2(IEnumerable<string> s)
        {
            var v = new List<string>(s);

            foreach(var vvv in s)
            {
                var t = DeleteComment(vvv);
                yield return t;
                if (t != vvv) yield break;
            }
        }

        public static IEnumerable<List<string>> LogicalRows(IEnumerable<string> s)
        {
            var v = new List<string>();

            foreach(var sss in s)
            {
                if (sss.Trim() == "") {
                    if (v.Count > 0)
                    {
                        yield return v;
                        v = new List<string>();
                    }
                    continue;
                }

                if(sss.Reverse().First()!='_')
                {
                    v.Add(sss);
                    yield return v;
                    v = new List<string>();
                    continue;
                }

                v.Add(sss);
            }

            if (v.Count > 0)
                yield return v;
        }
    }
}
