using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using EasyWord;
namespace WordTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择保存的路径";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string folder = dialog.SelectedPath;
                EasyWord.EasyWord word=new EasyWord.EasyWord(txtFileName.Text);

                word.CreateWord(folder);
                
            }
        }
    }
}
