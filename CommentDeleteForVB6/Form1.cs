using System;
using System.Text;
using System.Windows.Forms;
using System.IO;

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

            if (v.ShowDialog() == DialogResult.OK)
            {
                var vb6module = new VB6Source(v.FileName);
                File.WriteAllLines(v.FileName, vb6module.CommentDeleted,Encoding.Default);
            }

        }
    }
}
