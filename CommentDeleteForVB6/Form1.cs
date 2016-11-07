using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace CommentDeleteForVB6
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog v = new OpenFileDialog();

            v.Title = "Select VB6 Module File";

            if (v.ShowDialog() != DialogResult.OK) return;

            if (MessageBox.Show("Selected file will be overwritten!" + Environment.NewLine + "Are you OK?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
                return;

            DoDeleteComent(v.FileName);
        }

        private static void DoDeleteComent(string s)
        {
            var vb6module = new VB6Source(s);
            File.WriteAllLines(s, vb6module.CommentDeleted.Where(p => p.Trim() != ""), Encoding.Default);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog v = new FolderBrowserDialog();

            v.Description = "Select VB6 Module Folder";

            if (v.ShowDialog() != DialogResult.OK) return;

            if (MessageBox.Show("Selected file in folder will be overwritten!" + Environment.NewLine + "Are you OK?", "Caution", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
                return;

            var exts = new[] {"bas","frm","cls" };

            foreach (var ext in exts)
            {
                foreach (var f in Directory.GetFiles(v.SelectedPath, "*." + ext))
                {
                    DoDeleteComent(f);
                }

            }
        }

    }
}

