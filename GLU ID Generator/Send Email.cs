using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.IO;
using System.Data.SqlClient;
using System.Threading;

namespace GLU_ID_Generator
{
    public partial class Send_Email : Form
    {
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams handleparam = base.CreateParams;
                handleparam.ExStyle |= 0x02000000;
                return handleparam;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            this.DoubleBuffered = true;
        }
        public static void SetDoubleBuffered(System.Windows.Forms.Control c)
        {
            //Taxes: Remote Desktop Connection and painting
            //http://blogs.msdn.com/oldnewthing/archive/2006/01/03/508694.aspx
            if (System.Windows.Forms.SystemInformation.TerminalServerSession)
                return;

            System.Reflection.PropertyInfo aProp =
                  typeof(System.Windows.Forms.Control).GetProperty(
                        "DoubleBuffered",
                        System.Reflection.BindingFlags.NonPublic |
                        System.Reflection.BindingFlags.Instance);

            aProp.SetValue(c, true, null);
        }
        private string ssl;
        private string port;
        private string server;
        private string esender;
        private string epassword;
        private string esignature;
        public Send_Email()
        {
            InitializeComponent();
        }
        public string filename;
        private void Send_Email_Load(object sender, EventArgs e)
        {
            load();
            loaddgv();
            guna2ComboBox1.Text = filename;
           
        }
       
        public void load()
        {
            DateTime date = DateTime.Now;
            string fromdate = date.AddDays(-7).ToString("MMM dd");
            string todate;

            if (date.ToString("MMM") == date.AddDays(-7).ToString("MMM"))
            {
                todate = "";
                todate = date.ToString("dd");
            }
            else
            {
                todate = "";
                todate = date.ToString("MMM dd");
            }

            
            guna2TextBox1.Text = "Biometric Report ("+ fromdate.ToString()+" - "+ todate.ToString()+")";
            label8.Text = Properties.Settings.Default.ssl.ToString();
            label9.Text = Properties.Settings.Default.portnumber.ToString();
            label10.Text = Properties.Settings.Default.smtpserver.ToString();
            label11.Text = Properties.Settings.Default.esender.ToString();
            label12.Text = Properties.Settings.Default.attachment.ToString();


            ssl = Properties.Settings.Default.ssl.ToString();
            port = Properties.Settings.Default.portnumber.ToString();
            server = Properties.Settings.Default.smtpserver.ToString();
            esender = Properties.Settings.Default.esender.ToString();
            epassword = Properties.Settings.Default.esenderpass.ToString();
            esignature = Properties.Settings.Default.signature.ToString();

            try
            {
                guna2ComboBox1.Items.Clear();
                var files = Directory.EnumerateFiles(label12.Text, "*.*", SearchOption.AllDirectories)
                .Where(s => s.EndsWith(".pdf"));
                foreach (string filePath in files) guna2ComboBox1.Items.Add(Path.GetFileName(filePath));
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                guna2Button4.Visible = false;
                guna2Button2.Visible = false;
            }
        }


        private string cons = Properties.Settings.Default.dbcon;
        public void loaddgv()
        {
            using (SqlConnection codeMaterial = new SqlConnection(cons))
            {
                dataGridView1.Rows.Clear();
                DataTable dt = new DataTable();
                codeMaterial.Open();
                string list = "Select name,email from tblContacts order by Id ASC";
                SqlCommand command = new SqlCommand(list, codeMaterial);
                SqlDataReader reader = command.ExecuteReader();
                dt.Load(reader);
                foreach (DataRow row in dt.Rows)
                {
                    int a = dataGridView1.Rows.Add();
                    dataGridView1.Rows[a].Cells["ename"].Value = row["name"].ToString();
                    dataGridView1.Rows[a].Cells["email"].Value = row["email"].ToString();
                }


                //dataGridView1.DataSource = dt;
                codeMaterial.Close();
                dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                dataGridView1.RowHeadersVisible = false;
            }
        }
        private void send()
        {
            Cursor.Current = Cursors.WaitCursor;
           
                //Smpt Client Details
                //gmail >> smtp server : smtp.gmail.com, port : 587 , ssl required
                //yahoo >> smtp server : smtp.mail.yahoo.com, port : 587 , ssl required
                SmtpClient clientDetails = new SmtpClient();
                clientDetails.Port = Convert.ToInt32(port.Trim());
                clientDetails.Host = server.Trim();
                clientDetails.EnableSsl = Convert.ToBoolean(ssl);
                clientDetails.DeliveryMethod = SmtpDeliveryMethod.Network;
                clientDetails.UseDefaultCredentials = false;
                clientDetails.Credentials = new NetworkCredential(esender.Trim(), epassword.Trim());

                //Message Details

            foreach (DataGridViewRow rows in dataGridView1.Rows)
            {
                try
                {

                    MailMessage mailDetails = new MailMessage();
                    mailDetails.From = new MailAddress(esender.Trim());

                    dataGridView1.BeginInvoke((Action)delegate ()
                    {
                        rows.Cells["status"].Value = "Sending...";
                    });
                    if (guna2TextBox2.Text == "")
                    {

                        mailDetails.To.Add(rows.Cells["email"].Value.ToString().Trim());

                    }
                    else
                    {
                        mailDetails.To.Add(guna2TextBox2.Text.Trim());
                    }

                    mailDetails.Subject = guna2TextBox1.Text.Trim();
                    mailDetails.IsBodyHtml = true;
                    if (guna2TextBox3.Text == "")
                    {
                        mailDetails.Body = "Good day.." + "<br /><br />" + "Please see attached file for the " + guna2TextBox1.Text + "." + "<br /><br />" + "Thank you.." + "<br /><br />" + esignature;
                    }
                    else
                    {

                        //var textBoxText = guna2TextBox3.Text.Replace(Environment.NewLine , "<br />");
                        mailDetails.Body = guna2TextBox3.Text.Replace(Environment.NewLine, "<br />");
                    }
                    guna2ComboBox1.BeginInvoke((Action)delegate ()
                    {
                        if (guna2ComboBox1.Text != null || guna2ComboBox1.Text != "")
                        {
                            Attachment attachment = new Attachment(label12.Text + "\\" + guna2ComboBox1.Text);
                            mailDetails.Attachments.Add(attachment);
                        }
                    });
                    clientDetails.Send(mailDetails);

                    dataGridView1.BeginInvoke((Action)delegate ()
                    {
                        if (rows.Cells["status"].Value.ToString() != "Invalid")
                        {
                            rows.Cells["status"].Value = "Sent";
                        }
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    dataGridView1.BeginInvoke((Action)delegate ()
                    {
                        rows.Cells["status"].Value = "Invalid";
                        rows.Cells["status"].Style.ForeColor = Color.FromArgb(207, 123, 122);
                    });
                }
            }
            MessageBox.Show("Your mail has been sent.");
            Cursor.Current = Cursors.Default;
        }
        public void message()
        {

        }
        private void singlesend()
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                MailMessage mailDetails = new MailMessage();
                mailDetails.From = new MailAddress(esender.Trim());
                mailDetails.To.Add(guna2TextBox2.Text.Trim());
                mailDetails.Subject = guna2TextBox1.Text.Trim();
                mailDetails.IsBodyHtml = true;
                if (guna2TextBox3.Text == "")
                {
                    mailDetails.Body = "Good day.." + "<br /><br />" + "Please see attached file for the " + guna2TextBox1.Text + "." + "<br /><br />" + "Thank you.." + "<br /><br />" + esignature;
                }
                else
                //Add custom message for email body.
                {
                    mailDetails.Body = guna2TextBox3.Text.Replace(Environment.NewLine, "<br />");
                }
                if (guna2ComboBox1.Text != "")
                {
                    Attachment attachment = new Attachment(label12.Text + "\\" + guna2ComboBox1.Text);
                    mailDetails.Attachments.Add(attachment);
                }
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", Convert.ToInt32(port.Trim()));
                // smtp.Host = "smtp.gmail.com"; //Or Your SMTP Server Address
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = Convert.ToBoolean(ssl);
                smtp.Credentials = new System.Net.NetworkCredential(esender.Trim(), epassword.Trim());
               
                smtp.Send(mailDetails);

                MessageBox.Show("Your mail has been sent.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Cursor.Current = Cursors.Default;
        }
        private void guna2Button4_Click(object sender, EventArgs e)
        {
            if (guna2TextBox2.Text =="")
            {
                MessageBox.Show("Please enter an email address");
            }
            //else if (guna2ComboBox1.Text == "")
            //{
            //    MessageBox.Show("Please select a Biometric Report");
            //}
            else
            {
                DialogResult dialogResult1 = MessageBox.Show("Are you sure to send this???", "Send?", MessageBoxButtons.YesNo);
                if (dialogResult1 == DialogResult.Yes)
                {
                    singlesend();
                }
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            contacts c = new contacts();
            c.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            guna2TextBox2.Text = dataGridView1.CurrentRow.Cells["email"].Value.ToString();
        }
        Thread thread;
        private void guna2Button2_Click(object sender, EventArgs e)
        {
            guna2TextBox2.Text = "";
             if (guna2ComboBox1.Text == "")
            {
                MessageBox.Show("Please select a Biometric Report");
            }
            else
            {
                DialogResult dialogResult1 = MessageBox.Show("Are you sure to send this???", "Send?", MessageBoxButtons.YesNo);
                if (dialogResult1 == DialogResult.Yes)
                {
                    thread =
             new Thread(new ThreadStart(send));
                    thread.Start();
                    //send();
                }
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            esettings es = new esettings();
            es.ShowDialog();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            deletecontacts d = new deletecontacts();
            d.ShowDialog();
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            customMessage d = new customMessage();
            d.ShowDialog();
        }
    }
}
