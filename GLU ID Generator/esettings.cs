using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class esettings : Form
    {
        public esettings()
        {
            InitializeComponent();
        }

        private void esettings_Load(object sender, EventArgs e)
        {
            guna2TextBox1.Text = Properties.Settings.Default.esender.ToString();
            guna2TextBox2.Text = Properties.Settings.Default.esenderpass.ToString();
            guna2TextBox3.Text = Properties.Settings.Default.smtpserver.ToString();
            guna2TextBox4.Text = Properties.Settings.Default.portnumber.ToString();
            guna2ComboBox1.Text = Properties.Settings.Default.ssl.ToString();
            guna2TextBox5.Text = Properties.Settings.Default.attachment.ToString();
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This directory means is use for seaching attachments of all PDF file when sending emails.");
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Please enter you google account username and password to be used for sending emails."+Environment.NewLine+ Environment.NewLine +"The SMTP,PORT and SSL is from the Google Server. You can get it online. It is a gateway to send emails");
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                  

                    var path3 = fbd.SelectedPath.ToString();

                  
                        guna2TextBox5.Text = path3.ToString();
                }
            }
        }

        Send_Email obj = (Send_Email)Application.OpenForms["Send_Email"];
        private void guna2Button2_Click(object sender, EventArgs e)
        {
            //Properties.Settings.Default.biopath = path3;
            //Properties.Settings.Default.Save();

            Properties.Settings.Default.esender = guna2TextBox1.Text;
            Properties.Settings.Default.esenderpass = guna2TextBox2.Text;
            Properties.Settings.Default.smtpserver = guna2TextBox3.Text;
            Properties.Settings.Default.portnumber = guna2TextBox4.Text;
            Properties.Settings.Default.ssl = Convert.ToBoolean(guna2ComboBox1.Text);
            Properties.Settings.Default.attachment= guna2TextBox5.Text;
            Properties.Settings.Default.Save();
            obj.load();
            MessageBox.Show("Saved");
        }
    }
}
