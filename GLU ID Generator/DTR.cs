using MailKit.Security;
using Microsoft.Reporting.WinForms;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class DTR : Form
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

        private string cons = Properties.Settings.Default.dbcon;
        public DTR()
        {
            InitializeComponent();
        }

        private void DTR_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            guna2TextBox1.Text = Properties.Settings.Default.path;
            guna2TextBox2.Text = Properties.Settings.Default.biopath;
            this.reportViewer1.RefreshReport();
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            Selection w = new Selection();
            this.Hide();
            w.ShowDialog();
            this.Close();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {


            var path3 = guna2TextBox2.Text + "\\Biometric.xlsx";

            if (File.Exists(path3))
            {
                //string SourceConstr = "";
                //if (path3.CompareTo(".xls") == 0)
                //    SourceConstr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path3 + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                //else
                //    SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path3 + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                //                                                                                                                            ////using (OleDbConnection con = new OleDbConnection(conn))


                Cursor.Current = Cursors.WaitCursor;
                dataGridView1.Rows.Clear();
                DataTable dtExcel = new DataTable();
                string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path3 + ";Extended Properties='Excel 8.0;HDR=Yes';";
                OleDbConnection con = new OleDbConnection(SourceConstr);
                string query = "Select * from [Sheet1$]";
                OleDbDataAdapter data = new OleDbDataAdapter(query, con);
                data.Fill(dtExcel);
                foreach (DataRow row in dtExcel.Rows)
                {
                    int a = dataGridView1.Rows.Add();
                    dataGridView1.Rows[a].Cells["userid"].Value = row[0].ToString().Trim();
                    dataGridView1.Rows[a].Cells["employeeid"].Value = row[1].ToString().Trim();
                    dataGridView1.Rows[a].Cells["name"].Value = row[2].ToString().Trim();
                    dataGridView1.Rows[a].Cells["clocking"].Value = row[3].ToString().Trim();
                    dataGridView1.Rows[a].Cells["CheckType"].Value = row[4].ToString().Trim();
                }
                foreach (DataGridViewRow erows in dataGridView1.Rows)
                {
                    DateTime date = DateTime.Parse(erows.Cells["clocking"].Value.ToString());
                    string fdate = date.ToString("MM/dd/yyyy HH:mm");
                    erows.Cells["clocking"].Value = fdate.ToString();

                    if (erows.Cells["CheckType"].Value.ToString() == "0" || erows.Cells["CheckType"].Value.ToString() == "3" || erows.Cells["CheckType"].Value.ToString() == "4")
                    {
                        erows.Cells["CheckType"].Value = "Check-In";
                    }
                    else
                    {
                        erows.Cells["CheckType"].Value = "Check-Out";
                    }
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    int n = i + 1;
                    dataGridView1.Rows[i].Cells["Column1"].Value = n.ToString();
                }

                using (SqlConnection tblOTS = new SqlConnection(cons))
                {
                    string sqlTrunc = "TRUNCATE TABLE tblDTR";
                    SqlCommand cmd = new SqlCommand(sqlTrunc, tblOTS);
                    tblOTS.Open();
                    cmd.ExecuteNonQuery();
                    tblOTS.Close();
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        tblOTS.Open();
                        string insStmt2 = "insert into tblDTR ([userid],[employeeid],[Name],[Clocking],[CheckType],[id]) values" +
                                      " (@userid,@employeeid,@Name,@Clocking,@CheckType,@id)";

                        SqlCommand insCmd2 = new SqlCommand(insStmt2, tblOTS);

                        insCmd2.Parameters.AddWithValue("@userid", dataGridView1.Rows[i].Cells["userid"].Value.ToString());
                        insCmd2.Parameters.AddWithValue("@employeeid", dataGridView1.Rows[i].Cells["employeeid"].Value.ToString());
                        insCmd2.Parameters.AddWithValue("@Name", dataGridView1.Rows[i].Cells["Name"].Value.ToString());
                        insCmd2.Parameters.AddWithValue("@Clocking", dataGridView1.Rows[i].Cells["Clocking"].Value.ToString());
                        insCmd2.Parameters.AddWithValue("@CheckType", dataGridView1.Rows[i].Cells["CheckType"].Value.ToString());
                        insCmd2.Parameters.AddWithValue("@id", dataGridView1.Rows[i].Cells["Column1"].Value.ToString());
                        insCmd2.ExecuteNonQuery();
                        tblOTS.Close();



                        tblOTS.Open();
                        SqlCommand check_User_Name = new SqlCommand("SELECT COUNT(*) FROM [tblDTRHistory] WHERE employeeid = @employeeid AND Clocking = @Clocking", tblOTS);
                        check_User_Name.Parameters.AddWithValue("@employeeid", dataGridView1.Rows[i].Cells["employeeid"].Value.ToString());
                        check_User_Name.Parameters.AddWithValue("@Clocking", dataGridView1.Rows[i].Cells["Clocking"].Value.ToString());
                        int UserExist = (int)check_User_Name.ExecuteScalar();

                        if (UserExist == 0)
                        {
                            tblOTS.Close();

                            tblOTS.Open();
                            string insStmt3 = "insert into tblDTRHistory ([userid],[employeeid],[Name],[Clocking],[CheckType]) values" +
                                          " (@userid,@employeeid,@Name,@Clocking,@CheckType)";
                            SqlCommand insCmd3 = new SqlCommand(insStmt3, tblOTS);
                            insCmd3.Parameters.AddWithValue("@userid", dataGridView1.Rows[i].Cells["userid"].Value.ToString());
                            insCmd3.Parameters.AddWithValue("@employeeid", dataGridView1.Rows[i].Cells["employeeid"].Value.ToString());
                            insCmd3.Parameters.AddWithValue("@Name", dataGridView1.Rows[i].Cells["Name"].Value.ToString());
                            insCmd3.Parameters.AddWithValue("@Clocking", dataGridView1.Rows[i].Cells["Clocking"].Value.ToString());
                            insCmd3.Parameters.AddWithValue("@CheckType", dataGridView1.Rows[i].Cells["CheckType"].Value.ToString());

                            insCmd3.ExecuteNonQuery();
                            tblOTS.Close();

                        }else
                        {
                            tblOTS.Close();
                        }
                    }
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {


                    }
                }
                string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\DTRVIEW.rdlc";

                reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
                ReportDataSource dr = new ReportDataSource("DataSet1", WORKERSDATA());
                this.reportViewer1.LocalReport.ReportPath = exeFolder;
                this.reportViewer1.LocalReport.DataSources.Add(dr);
                this.reportViewer1.RefreshReport();
                var deviceInfo = @"<DeviceInfo>
                    <EmbedFonts>None</EmbedFonts>
                   </DeviceInfo>";

                //byte[] bytes = rdlc.Render("PDF", deviceInfo);

                byte[] Bytes = reportViewer1.LocalReport.Render(format: "PDF", deviceInfo);
                DateTime fadate = DateTime.Now;
                string ffdate = fadate.ToString("MM-dd-yyyy");
                using (FileStream stream = new FileStream(guna2TextBox1.Text + "\\Biometric " + ffdate.ToString() + ".pdf", FileMode.Create))
                {
                    stream.Write(Bytes, 0, Bytes.Length);
                }
                filename = "Biometric " + ffdate.ToString() + ".pdf";
                Cursor.Current = Cursors.Default;
            }
            else
            {
                MessageBox.Show("The PERID.DBF doesn't exist make sure to locate the file with the specified folder.");
            }
        }
        string filename;
        private void guna2Button1_Click(object sender, EventArgs e)
        {

        }
        private DataTable WORKERSDATA()
        {
            using (SqlConnection dbDR = new SqlConnection(cons))
            {
                DataTable dt = new DataTable();
                dbDR.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM tblDTR ORDER BY Name,Clocking ASC", dbDR);
                SqlDataReader rd = cmd.ExecuteReader();
                dt.Load(rd);
                dbDR.Close();
                return dt;
            }
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void guna2Button3_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    var path3 = fbd.SelectedPath.ToString();
                    Cursor.Current = Cursors.WaitCursor;
                    Properties.Settings.Default.path = path3;
                    Properties.Settings.Default.Save();
                    guna2TextBox1.Text = Properties.Settings.Default.path;
                    Cursor.Current = Cursors.Default;
                }
            }
        }
        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();
                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    var path3 = fbd.SelectedPath.ToString();
                    Cursor.Current = Cursors.WaitCursor;
                    Properties.Settings.Default.biopath = path3;
                    Properties.Settings.Default.Save();
                    guna2TextBox2.Text = Properties.Settings.Default.biopath;
                    Cursor.Current = Cursors.Default;
                }
            }
        }
        private void guna2Button4_Click(object sender, EventArgs e)
        {
            //SmtpClient mailServer = new SmtpClient("smtp.gmail.com");
            //mailServer.Port = 587;
            //mailServer.UseDefaultCredentials = false;

            //mailServer.Credentials = new System.Net.NetworkCredential("jedzerna4@gmail.com", "jlcjed1998");
            //mailServer.EnableSsl = true;
            //string from = "jedzerna4@gmail.com";
            //string to = "rinajoybalioggujc@gmail.com";
            //MailMessage msg = new MailMessage(from, to);
            //msg.Subject = "Example Subject";
            //msg.Body = "The message goes here.";
            //DateTime fadate = DateTime.Now;
            //string ffdate = fadate.ToString("MM-dd-yyyy");

            //msg.Attachments.Add(new Attachment(guna2TextBox1.Text + "\\Biometric " + ffdate.ToString() + ".pdf"));

            //mailServer.Send(msg);
            //try
            //{
            //    MailMessage message = new MailMessage();
            //    SmtpClient smtp = new SmtpClient();

            //    message.From = new MailAddress("jedzerna4@gmail.com");
            //    message.To.Add(new MailAddress("rinajoybalioggujc@gmail.com"));
            //    message.Subject = "Test";
            //    message.Body = "Content";

            //    smtp.Port = 587;
            //    smtp.Host = "smtp.gmail.com";
            //    smtp.EnableSsl = true;
            //    smtp.UseDefaultCredentials = false;
            //    smtp.Credentials = new NetworkCredential("jedzerna4@gmail.com", "jlcjed1998");
            //    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    smtp.Send(message);
            //    MessageBox.Show("mail Send");

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("err: " + ex.Message);
            //}
            Send_Email s = new Send_Email();
            s.filename = filename;
            s.ShowDialog();
        }


        private void guna2Button5_Click(object sender, EventArgs e)
        {

        }

        private void DTR_FormClosed(object sender, FormClosedEventArgs e)
        {
            Selection d = new Selection();
            this.Hide();
            d.ShowDialog();
            this.Close();
        }

        private void guna2Button5_Click_1(object sender, EventArgs e)
        {
            customSort s = new customSort();
            s.ShowDialog();
        }
    }
}
