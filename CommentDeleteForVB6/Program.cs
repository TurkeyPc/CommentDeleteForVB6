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
    }
}
