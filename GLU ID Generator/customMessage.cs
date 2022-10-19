using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class customMessage : Form
    {
        public customMessage()
        {
            InitializeComponent();
        }

        private void customMessage_Load(object sender, EventArgs e)
        { 
            guna2TextBox1.Text = Properties.Settings.Default.signature;
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            guna2HtmlLabel1.Text = guna2TextBox1.Text;
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.signature = guna2TextBox1.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Saved"); 
        }
    }
}