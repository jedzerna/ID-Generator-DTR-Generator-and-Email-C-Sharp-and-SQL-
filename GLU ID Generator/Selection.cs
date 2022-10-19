using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class Selection : Form
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
        public Selection()
        {
            InitializeComponent();
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            check();
            if (val == true)
            {
                WorkersID w = new WorkersID();
                this.Hide();
                w.ShowDialog();
                this.Close();
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            check();
            if (val == true)
            {
                EmployeeForm w = new EmployeeForm();
                this.Hide();
                w.ShowDialog();
                this.Close();
            }
        }

        private void Selection_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        public const int WM_NCLBUTTONDOWN = 0x00A1;
        public const int HT_CAPTION = 0x2;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            //System.Threading.Thread thread =
            //       new System.Threading.Thread(new System.Threading.ThreadStart(check));
            //thread.Start();
            check();
            if (val == true)
            {
                this.Hide();
                //new Thread(() => new DTR().ShowDialog()).Start();
                DTR d = new DTR();
                d.ShowDialog();
                this.Close();
            }
        }
        private void Selection_Load(object sender, EventArgs e)
        {
            check();
            guna2Panel1.SendToBack();
        }
        private string con = Properties.Settings.Default.dbcon;
        private bool val;
        private void check()
        {
            LaunchCommandLineApp();
            try
            {
                string provider = "System.Data.SqlClient"; // for example
                DbProviderFactory factory = DbProviderFactories.GetFactory(provider);
                using (DbConnection conn = factory.CreateConnection())
                {
                    conn.ConnectionString = Properties.Settings.Default.dbcon;
                    conn.Open();
                    val = true;
                }
            }
            catch (Exception ex)
            {
                val = false;
                MessageBox.Show("Can't find DBEmployee.mdf."+Environment.NewLine + Environment.NewLine + Environment.NewLine+"Error:" + ex.Message);
             
            }
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }
        static void LaunchCommandLineApp()
        {
            // For the example
            //const string ex1 = @"C:\Users\Jed Zerna\source\repos\Pay Slip\Pay Slip\bin\Debug";

            //// Use ProcessStartInfo class
            //ProcessStartInfo startInfo = new ProcessStartInfo();
            //startInfo.CreateNoWindow = false;
            //startInfo.UseShellExecute = false;
            //startInfo.FileName = "Pay Slip.exe";
            //startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            //startInfo.Arguments = "-f j -o \"" + ex1;

            //try
            //{
            //    Process.Start(@"" + Properties.Settings.Default.pspath + "\\Pay Slip.exe");
            //    this.Close();
                
            //}
            //catch(Exception ex)
            //{
            //    // Log error.
            //    MessageBox.Show(ex.Message);
            //    using (var fbd = new FolderBrowserDialog())
            //    {
            //        DialogResult result = fbd.ShowDialog();

            //        if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            //        {

            //            var path3 = fbd.SelectedPath.ToString();
            //            Properties.Settings.Default.pspath = path3.ToString();
            //            Properties.Settings.Default.Save();
            //        }
            //    }
            //}
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"" + Properties.Settings.Default.pspath + "\\Pay Slip.exe");
                this.Close();

            }
            catch (Exception ex)
            {
                // Log error.
                MessageBox.Show(ex.Message);
                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {

                        var path3 = fbd.SelectedPath.ToString();
                        Properties.Settings.Default.pspath = path3.ToString();
                        Properties.Settings.Default.Save();
                    }
                }
            }
        }

        private void Selection_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void Selection_SizeChanged(object sender, EventArgs e)
        {
           
        }

        private void Selection_Resize(object sender, EventArgs e)
        {
          
        }

        private void Selection_Move(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
       
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Browse MDF Files",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "mdf",
                Filter = "MDF files (*.mdf)|*.mdf",
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string con = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename="+ openFileDialog1.FileName + ";Integrated Security=True";

                Properties.Settings.Default.dbcon = con;
                Properties.Settings.Default.Save();

            }
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Properties.Settings.Default.dbcon);
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
         
        }
    }
}
