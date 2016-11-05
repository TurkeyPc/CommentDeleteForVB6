using System;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Linq;

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

            var vb6module = new VB6Source(v.FileName);
            File.WriteAllLines(v.FileName, vb6module.CommentDeleted.Where(p => p.Trim() != ""), Encoding.Default);
        }

    }
}

