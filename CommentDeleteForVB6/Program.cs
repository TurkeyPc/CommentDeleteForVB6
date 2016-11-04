using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CommentDeleteForVB6
{
    static class Program
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

        static string DeleteComment(string s)
        {

        }
    }
}
