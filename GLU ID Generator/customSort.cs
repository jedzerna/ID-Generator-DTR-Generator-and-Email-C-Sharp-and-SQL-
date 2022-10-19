using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class customSort : Form
    {
        public customSort()
        {
            InitializeComponent();
        }
        string tb1;
        string tb2;
        private void customSort_Load(object sender, EventArgs e)
        {
            tb1 = Properties.Settings.Default.path;
            tb2 = Properties.Settings.Default.biopath;
            this.reportViewer1.RefreshReport();
            maskedTextBox4.Text = DateTime.Now.AddDays(1).ToString("MM/dd/yyyy");
            maskedTextBox2.Text = DateTime.Now.AddDays(1).ToString("MM/dd/yyyy");
            maskedTextBox1.Text = DateTime.Now.AddMonths(-1).ToString("MM/dd/yyyy");
            maskedTextBox5.Text = DateTime.Now.AddMonths(-1).ToString("MM/dd/yyyy");
        }

        private void guna2RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2RadioButton2.Checked)
            {
                date.Visible = true;
            }
            else
            {
                date.Visible = false;
            }
        }

        private void guna2RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2RadioButton1.Checked)
            {
                emp.Visible = true;
            }
            else
            {
                emp.Visible = false;
            }
        }

        private string cons = Properties.Settings.Default.dbcon;
        private DataTable WORKERSDATA()
        {
            using (SqlConnection dbDR = new SqlConnection(cons))
            {
                DataTable dt = new DataTable();
                dt.Rows.Clear();
                dbDR.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM tblDTRHistory WHERE Clocking BETWEEN @date1 AND @date2 ORDER BY Name,Clocking ASC", dbDR);
                cmd.Parameters.AddWithValue("@date1", DateTime.Parse(maskedTextBox1.Text));
                cmd.Parameters.AddWithValue("@date2", DateTime.Parse(maskedTextBox2.Text));
                SqlDataReader rd = cmd.ExecuteReader();
                dt.Load(rd);
                dbDR.Close();
                return dt;
            }
        }
        private DataTable WORKERSDATA2()
        {
            using (SqlConnection dbDR = new SqlConnection(cons))
            {
                DataTable dt = new DataTable();
                dt.Rows.Clear();
                dbDR.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM tblDTRHistory WHERE Clocking BETWEEN @date1 AND @date2 AND employeeid = @employeeid ORDER BY Name,Clocking ASC", dbDR);
                cmd.Parameters.AddWithValue("@date1", DateTime.Parse(maskedTextBox5.Text));
                cmd.Parameters.AddWithValue("@date2", DateTime.Parse(maskedTextBox4.Text));
                cmd.Parameters.AddWithValue("@employeeid", guna2TextBox1.Text);
                SqlDataReader rd = cmd.ExecuteReader();
                dt.Load(rd);
                dbDR.Close();
                return dt;
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            DateTime temp;
            if (DateTime.TryParse(maskedTextBox1.Text, out temp))
            {
                // Yay :)
            }
            else
            {
                MessageBox.Show("Invalid Date");
                DateTime d = DateTime.Now;
                maskedTextBox1.Text = d.ToString("MM/dd/yyyy");
                return;
            }
            DateTime temp1;
            if (DateTime.TryParse(maskedTextBox2.Text, out temp1))
            {
                // Yay :)
            }
            else
            {
                MessageBox.Show("Invalid Date");
                DateTime d = DateTime.Now;
                maskedTextBox2.Text = d.ToString("MM/dd/yyyy");
                return;
            }
            if (thread != null)
            {
                // if running, abort it.  
                thread.Abort();
            }
            this.reportViewer1.Reset();
            thread =
            new Thread(new ThreadStart(start1));
            thread.Start();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            DateTime temp;
            if (DateTime.TryParse(maskedTextBox5.Text, out temp))
            {
                // Yay :)
            }
            else
            {
                MessageBox.Show("Invalid Date");
                DateTime d = DateTime.Now;
                maskedTextBox5.Text = d.ToString("MM/dd/yyyy");
                return;
            }
            DateTime temp1;
            if (DateTime.TryParse(maskedTextBox4.Text, out temp1))
            {
                // Yay :)
            }
            else
            {
                MessageBox.Show("Invalid Date");
                DateTime d = DateTime.Now;
                maskedTextBox4.Text = d.ToString("MM/dd/yyyy");
                return;
            }
            if (guna2TextBox1.Text == "")
            {
                MessageBox.Show("Please enter an EMPCODE");
                return;
            }
            if (thread != null)
            {
                // if running, abort it.  
                thread.Abort();
            }
            this.reportViewer1.Reset();
            thread =
            new Thread(new ThreadStart(start2));
            thread.Start();
        }
        Thread thread;
        private void start1()
        {
            guna2ProgressIndicator1.BeginInvoke((Action)delegate ()
            {
                guna2ProgressIndicator1.Start();
                guna2ProgressIndicator1.Visible = true;
            });
            reportViewer1.BeginInvoke((Action)delegate ()
            {
                reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
            });
            Thread.Sleep(1000);

            string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\DTRVIEW.rdlc";
            ReportDataSource dr = new ReportDataSource("DataSet1", WORKERSDATA());
            this.reportViewer1.LocalReport.ReportPath = exeFolder;
            this.reportViewer1.LocalReport.DataSources.Add(dr);
            reportViewer1.BeginInvoke((Action)delegate ()
            {
                this.reportViewer1.RefreshReport();
            });
            Thread.Sleep(1000);
            var deviceInfo = @"<DeviceInfo>
                    <EmbedFonts>None</EmbedFonts>
                   </DeviceInfo>";

            //byte[] bytes = rdlc.Render("PDF", deviceInfo);

            byte[] Bytes = reportViewer1.LocalReport.Render(format: "PDF", deviceInfo);
            DateTime fadate = DateTime.Now;
            string ffdate = fadate.ToString("MM-dd-yyyy");
            using (FileStream stream = new FileStream(tb1 + "\\Biometric " + ffdate.ToString() + ".pdf", FileMode.Create))
            {
                stream.Write(Bytes, 0, Bytes.Length);
            }
            //filename = "Biometric " + ffdate.ToString() + ".pdf";
            Cursor.Current = Cursors.Default;
            guna2ProgressIndicator1.BeginInvoke((Action)delegate ()
            {
                guna2ProgressIndicator1.Stop();
                guna2ProgressIndicator1.Visible = false;
            });
            Thread.Sleep(1000);
            thread.Abort();
        }
        private void start2()
        {
            guna2ProgressIndicator1.BeginInvoke((Action)delegate ()
            {
                guna2ProgressIndicator1.Start();
                guna2ProgressIndicator1.Visible = true;
            });
            Thread.Sleep(1000);
            reportViewer1.BeginInvoke((Action)delegate ()
            {
            reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
            });
            Thread.Sleep(1000);
            string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\DTRVIEW.rdlc";
            ReportDataSource dr = new ReportDataSource("DataSet1", WORKERSDATA2());
            this.reportViewer1.LocalReport.ReportPath = exeFolder;
            this.reportViewer1.LocalReport.DataSources.Add(dr);

            reportViewer1.BeginInvoke((Action)delegate ()
            {
                this.reportViewer1.RefreshReport();
            });
            var deviceInfo = @"<DeviceInfo>
                    <EmbedFonts>None</EmbedFonts>
                   </DeviceInfo>";

            //byte[] bytes = rdlc.Render("PDF", deviceInfo);
            byte[] Bytes = reportViewer1.LocalReport.Render(format: "PDF", deviceInfo);
            DateTime fadate = DateTime.Now;
            string ffdate = fadate.ToString("MM-dd-yyyy");
            using (FileStream stream = new FileStream(tb1 + "\\Biometric " + ffdate.ToString() + ".pdf", FileMode.Create))
            {
                stream.Write(Bytes, 0, Bytes.Length);
            }
            //filename = "Biometric " + ffdate.ToString() + ".pdf";
            Cursor.Current = Cursors.Default;

            guna2ProgressIndicator1.BeginInvoke((Action)delegate ()
            {
                guna2ProgressIndicator1.Stop();
                guna2ProgressIndicator1.Visible = false;
            });
            Thread.Sleep(1000);
            thread.Abort();
        }

        private void customSort_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (thread != null)
            { 
                thread.Abort();
            }
        }
    }
}
