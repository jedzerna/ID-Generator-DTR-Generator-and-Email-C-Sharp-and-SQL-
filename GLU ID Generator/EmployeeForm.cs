using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using ZXing.Common;
using ZXing;
using ZXing.QrCode;
using System.Data.OleDb;

namespace GLU_ID_Generator
{
    public partial class EmployeeForm : Form
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


        private string con = Properties.Settings.Default.dbcon; 
        public EmployeeForm()
        {
            InitializeComponent();
        }

        private void EmployeeForm_Load(object sender, EventArgs e)
        {
            load();
            loadhist();

        }
        DataTable employeename = new DataTable();
        string lb = "{";
        string rb = "}";
        public void load()
        {
            using (SqlConnection tblOTS = new SqlConnection(con))
            {
                employeename.Columns.Clear();
                employeename.Rows.Clear();
                employeename.Columns.Add("EMPCODE");
                employeename.Columns.Add("NAME");
                DataRow drLocal = null;
                using (SqlConnection dbDR = new SqlConnection(con))
                {
                    dbDR.Open();
                    DataTable dt = new DataTable();
                    string list = "SELECT EMPCODE,FIRST,MIDDLE,FAMILY FROM PERID";
                    SqlDataAdapter command = new SqlDataAdapter(list, dbDR);
                    command.Fill(dt);
                    foreach (DataRow item in dt.Rows)
                    {

                        int c = dataGridView1.Rows.Add();
                        string first = string.Concat(item["FIRST"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
                        string middle = string.Concat(item["MIDDLE"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
                        string family = string.Concat(item["FAMILY"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

                        
                        
                            drLocal = employeename.NewRow();
                            drLocal["EMPCODE"] = item["EMPCODE"].ToString();
                            drLocal["NAME"] = family +", "+first + " " + middle;
                            employeename.Rows.Add(drLocal);
                        
                    }
                    dbDR.Close();
                    dataGridView1.DataSource = employeename;
                }

            }
        }

        public void loadhist()
        {

            DataTable employeename = new DataTable();
            employeename.Columns.Clear();
            employeename.Rows.Clear();
            employeename.Columns.Add("Id");
            employeename.Columns.Add("EMPCODE");
            employeename.Columns.Add("NAME");
            employeename.Columns.Add("DATE");
            DataRow drLocal = null;
            using (SqlConnection dbDR = new SqlConnection(con))
            {
                dbDR.Open();
                DataTable dt = new DataTable();
                string list = "SELECT Id,idno,fname,mname,lname,date FROM tblEmpHist";
                SqlDataAdapter command = new SqlDataAdapter(list, dbDR);
                command.Fill(dt);
                foreach (DataRow item in dt.Rows)
                {

                    //int c = dataGridView1.Rows.Add();
                    string first = string.Concat(item["fname"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
                    string middle = string.Concat(item["mname"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
                    string family = string.Concat(item["lname"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

                    drLocal = employeename.NewRow();
                    drLocal["Id"] = item["Id"].ToString();
                    drLocal["EMPCODE"] = item["idno"].ToString();
                    drLocal["NAME"] = family + ", " + first + " " + middle;
                    drLocal["DATE"] = item["date"].ToString();
                    employeename.Rows.Add(drLocal);
                }
                dbDR.Close();
                employeename.DefaultView.Sort = "Id DESC";
                dataGridView2.DataSource = employeename;
            }

        }
        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            Selection w = new Selection();
            this.Hide();
            w.ShowDialog();
            this.Close();
        }
        private void guna2Button3_Click(object sender, EventArgs e)
        {
            generateqr();
        }
        private void generateqr()
        {
        
            var options = new QrCodeEncodingOptions
            {
                DisableECI = true,
                CharacterSet = "UTF-8",
                Width = 3000,
                Height = 3000,
            };
            var writer = new BarcodeWriter();
            writer.Format = BarcodeFormat.QR_CODE;
            writer.Options = options;

            var text = $"BEGIN:VCARD\n" +
                $"VERSION: 3.0\n" +
                //$"N: {guna2TextBox2.Text};;;\n" +
                $"N: {guna2TextBox12.Text}; {guna2TextBox2.Text};\n" +
                $"FN:  {guna2TextBox12.Text} {guna2TextBox2.Text}\n" +
                $"TITLE:{guna2TextBox3.Text} of G. Uymatiao Jr. Const.\n" +
                $"ORG:For Emergency Purposes\n" +
                $"URL; WORK: http://www.facebook.com/glujr \n" +
                $"TEL; TYPE = WORK:(035) 225 4716\n" +
                $"TEL; TYPE = VOICE: {guna2TextBox11.Text}\n" +
                $"TEL;CELL: {guna2TextBox11.Text}\n" +
                $"TEL;TYPE = CELL: {guna2TextBox11.Text}\n" +
                $"TEL; TYPE = FAX:(035)225 3333\n" +
                $"EMAIL; INTERNET: {guna2TextBox9.Text}\n" +
                $"ADR; INTL; PARCEL; WORK; CHARSET = utf - 8:; ;  {guna2TextBox6.Text}; {guna2TextBox8.Text}; ; ; PHIL;\n" +
                $"NOTE; CHARSET = utf - 8:For Emergency Contact Person ( Name:{guna2TextBox4.Text} \n Contact No:{guna2TextBox5.Text} \n Address:{guna2TextBox6.Text}, {guna2TextBox8.Text})  \n" +
                $"END:VCARD";

            if (String.IsNullOrWhiteSpace(text) || String.IsNullOrEmpty(text))
            {
                pictureBox1.Image = null;
                MessageBox.Show("Text not found", "Oops!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var qr = new ZXing.BarcodeWriter();
                qr.Options = options;
                qr.Format = ZXing.BarcodeFormat.QR_CODE;
                var result = new Bitmap(qr.Write(text));
                pictureBox2.Image = result;
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            // image filters  
            open.Filter = "Image Files(*.jpg; *.jpeg; *.png;)|*.jpg; *.jpeg; *.png;";
            if (open.ShowDialog() == DialogResult.OK)
            {
                // display image in picture box  
                pictureBox1.Image = new Bitmap(open.FileName);
                // image file path  
                //textBox1.Text = open.FileName;
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);

                    var path3 = fbd.SelectedPath.ToString() + "\\PERID.DBF";

                    if (File.Exists(path3))
                    {
                        using (SqlConnection tblOTS = new SqlConnection(con))
                        {
                            string sqlTrunc = "TRUNCATE TABLE PERID";
                            SqlCommand cmd = new SqlCommand(sqlTrunc, tblOTS);
                            tblOTS.Open();
                            cmd.ExecuteNonQuery();
                            tblOTS.Close();

                            OleDbConnection oConn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + fbd.SelectedPath.ToString() + "';Extended Properties=dBase IV;");
                            OleDbCommand command = new OleDbCommand("select * from PERID.DBF", oConn);
                            DataTable dt = new DataTable();
                            oConn.Open();
                            dt.Load(command.ExecuteReader());
                            oConn.Close();


                            DataTableReader reader = dt.CreateDataReader();
                            tblOTS.Open();  ///this is my connection to the sql server
                            SqlBulkCopy sqlcpy = new SqlBulkCopy(tblOTS);
                            sqlcpy.DestinationTableName = "PERID";  //copy the datatable to the sql table
                            sqlcpy.WriteToServer(dt);
                            tblOTS.Close();
                            reader.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("The PERID.DBF doesn't exist make sure to locate the file with the specified folder.");
                    }
                }
            }
            
        }
        private void saveToDatabase()
        {
            if (pictureBox1.Image == null )
            {
                MessageBox.Show("Please add an Image of the ID..");
            }
            else if(pictureBox2.Image == null)
            {
                MessageBox.Show("Please Generate QR Code...");
            }
            else
            {
                using (SqlConnection itemCode = new SqlConnection(con))
                {

                    MemoryStream ms = new MemoryStream();
                    pictureBox1.Image.Save(ms, pictureBox1.Image.RawFormat);
                    byte[] pic1 = ms.ToArray();


                    MemoryStream ms1 = new MemoryStream();
                    pictureBox2.Image.Save(ms1, System.Drawing.Imaging.ImageFormat.Png);
                    byte[] pic2 = ms1.ToArray();

                    itemCode.Open();

                    SqlCommand cmd = new SqlCommand("update tblEmp set name=@name,designation=@designation,idno=@idno,ename=@ename,enumber=@enumber,estreet=@estreet,ecity=@ecity,sss=@sss,tin=@tin,image=@image,qrcode=@qrcode where Id=@Id", itemCode);

                    cmd.Parameters.AddWithValue("@Id", "1");
                    cmd.Parameters.AddWithValue("@name", guna2TextBox12.Text + ", " + guna2TextBox2.Text + " " + guna2TextBox13.Text);
                    cmd.Parameters.AddWithValue("@designation", guna2TextBox3.Text);
                    cmd.Parameters.AddWithValue("@idno", guna2TextBox1.Text);
                    cmd.Parameters.AddWithValue("@ename", guna2TextBox4.Text);
                    cmd.Parameters.AddWithValue("@enumber", guna2TextBox5.Text);
                    cmd.Parameters.AddWithValue("@estreet", guna2TextBox6.Text);
                    cmd.Parameters.AddWithValue("@ecity", guna2TextBox8.Text);
                    cmd.Parameters.AddWithValue("@sss", guna2TextBox7.Text);
                    cmd.Parameters.AddWithValue("@tin", guna2TextBox10.Text);
                    cmd.Parameters.AddWithValue("@image", SqlDbType.Image).Value = pic1;
                    cmd.Parameters.AddWithValue("@qrcode", SqlDbType.Image).Value = pic2;
                    cmd.ExecuteNonQuery();

                    itemCode.Close();

                }
                using (SqlConnection itemCode = new SqlConnection(con))
                {

                    MemoryStream ms = new MemoryStream();
                    pictureBox1.Image.Save(ms, pictureBox1.Image.RawFormat);
                    byte[] pic1 = ms.ToArray();


                    MemoryStream ms1 = new MemoryStream();
                    pictureBox2.Image.Save(ms1, System.Drawing.Imaging.ImageFormat.Png);
                    byte[] pic2 = ms1.ToArray();

                    itemCode.Open();

                    SqlCommand cmd = new SqlCommand("insert into tblEmpHist ([fname],[designation],[idno],[ename],[enumber],[estreet],[ecity],[sss],[tin],[image],[qrcode],[date],[mname],[lname],[email],[number]) values (@fname,@designation,@idno,@ename,@enumber,@estreet,@ecity,@sss,@tin,@image,@qrcode,@date,@mname,@lname,@email,@number)", itemCode);

                    cmd.Parameters.AddWithValue("@fname", guna2TextBox2.Text);
                    cmd.Parameters.AddWithValue("@designation", guna2TextBox3.Text);
                    cmd.Parameters.AddWithValue("@idno", guna2TextBox1.Text);
                    cmd.Parameters.AddWithValue("@ename", guna2TextBox4.Text);
                    cmd.Parameters.AddWithValue("@enumber", guna2TextBox5.Text);
                    cmd.Parameters.AddWithValue("@estreet", guna2TextBox6.Text);
                    cmd.Parameters.AddWithValue("@ecity", guna2TextBox8.Text);
                    cmd.Parameters.AddWithValue("@sss", guna2TextBox7.Text);
                    cmd.Parameters.AddWithValue("@tin", guna2TextBox10.Text);
                    cmd.Parameters.AddWithValue("@image", SqlDbType.Image).Value = pic1;
                    cmd.Parameters.AddWithValue("@qrcode", SqlDbType.Image).Value = pic2;

                    DateTime dat = DateTime.Now;
                    string fdate = dat.ToString("MM/dd/yyyy HH:mm");
                    cmd.Parameters.AddWithValue("@date", fdate);
                    cmd.Parameters.AddWithValue("@mname", guna2TextBox13.Text);
                    cmd.Parameters.AddWithValue("@lname", guna2TextBox12.Text);
                    cmd.Parameters.AddWithValue("@email", guna2TextBox9.Text);
                    cmd.Parameters.AddWithValue("@number", guna2TextBox11.Text);
                    cmd.ExecuteNonQuery();

                    itemCode.Close();

                }
                loadhist();
            }
        }

        private void guna2CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox2.Checked == true)
            {
                guna2CheckBox1.Checked = false;
            }
            else
            {
                guna2CheckBox1.Checked = true;
            }
        }

        private void guna2CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox1.Checked == true)
            {
                guna2CheckBox2.Checked = false;
            }
            else
            {
                guna2CheckBox2.Checked = true;
            }
        }

        private void guna2TextBox14_TextChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox1.Checked == true)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("EMPCODE LIKE '%{0}%'", guna2TextBox14.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));

            }
            else if (guna2CheckBox2.Checked == true)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("NAME LIKE '%{0}%'", guna2TextBox14.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
            }
            else
            {
                MessageBox.Show("Please select what do you want to search...");
            }
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {
         
        }

        private void guna2TextBox12_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void guna2TextBox13_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void guna2TextBox3_TextChanged(object sender, EventArgs e)
        {
         
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("EMPCODE LIKE '%{0}%'", guna2TextBox1.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
        }

        private void guna2TextBox9_TextChanged(object sender, EventArgs e)
        {
         
        }

        private void guna2TextBox11_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void guna2TextBox7_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void guna2TextBox10_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void guna2TextBox4_TextChanged(object sender, EventArgs e)
        {
        
        }

        private void guna2TextBox5_TextChanged(object sender, EventArgs e)
        {
          
        }

        private void guna2TextBox6_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void guna2TextBox8_TextChanged(object sender, EventArgs e)
        {
      
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            using (SqlConnection tblSupplier = new SqlConnection(con))
            {
                tblSupplier.Open();
                DataTable dt = new DataTable();
                String query = "SELECT * FROM PERID WHERE EMPCODE = '" + dataGridView1.CurrentRow.Cells[0].Value + "'";
                SqlCommand cmd = new SqlCommand(query, tblSupplier);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    guna2TextBox1.Text = (rdr["EMPCODE"].ToString());
                    guna2TextBox2.Text = (rdr["FIRST"].ToString());
                    guna2TextBox13.Text = (rdr["MIDDLE"].ToString());
                    guna2TextBox12.Text = (rdr["FAMILY"].ToString());
                    guna2TextBox7.Text = (rdr["SSS"].ToString());
                    guna2TextBox6.Text = (rdr["ADD1"].ToString());
                    guna2TextBox8.Text = (rdr["ADD2"].ToString());
                }
                tblSupplier.Close();
            }
        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void EmployeeForm_MouseDown(object sender, MouseEventArgs e)
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

        private void EmployeeForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Selection d = new Selection();
            this.Hide();
            d.ShowDialog();
            this.Close();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            generateqr();
            saveToDatabase();
            Form1 f = new Form1();
            f.idtype = "GLU";
            f.ShowDialog();
        }
        private void guna2Button5_Click(object sender, EventArgs e)
        {
            generateqr();
            saveToDatabase();
            Form1 f = new Form1();
            f.idtype = "glubros";
            f.ShowDialog();
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            editing b = new editing();
            b.form = "employee";
            b.ShowDialog();
        }

        private void guna2CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox4.Checked == true)
            {
                guna2CheckBox3.Checked = false;
            }
            else
            {
                guna2CheckBox3.Checked = true;
            }
        }

        private void guna2CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox3.Checked == true)
            {
                guna2CheckBox4.Checked = false;
            }
            else
            {
                guna2CheckBox4.Checked = true;
            }
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            using (SqlConnection tblSupplier = new SqlConnection(con))
            {
                tblSupplier.Open();
                DataTable dt = new DataTable();
                String query = "SELECT * FROM tblEmpHist WHERE Id = '" + dataGridView2.CurrentRow.Cells["Id"].Value.ToString() + "'";
                SqlCommand cmd = new SqlCommand(query, tblSupplier);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    guna2TextBox2.Text = (rdr["fname"].ToString());
                    guna2TextBox3.Text = (rdr["designation"].ToString());
                    guna2TextBox1.Text = (rdr["idno"].ToString());
                    guna2TextBox4.Text = (rdr["ename"].ToString());
                    guna2TextBox5.Text = (rdr["enumber"].ToString());
                    guna2TextBox6.Text = (rdr["estreet"].ToString());
                    guna2TextBox8.Text = (rdr["ecity"].ToString());
                    guna2TextBox7.Text = (rdr["sss"].ToString());
                    guna2TextBox10.Text = (rdr["tin"].ToString());

                    if (rdr["image"] != DBNull.Value)
                    {
                        byte[] img = (byte[])(rdr["image"]);
                        MemoryStream mstream = new MemoryStream(img);
                        pictureBox1.Image = System.Drawing.Image.FromStream(mstream);
                    }
                    if (rdr["qrcode"] != DBNull.Value)
                    {
                        byte[] img = (byte[])(rdr["qrcode"]);
                        MemoryStream mstream = new MemoryStream(img);
                        pictureBox2.Image = System.Drawing.Image.FromStream(mstream);
                    }
                    guna2TextBox13.Text = (rdr["mname"].ToString());
                    guna2TextBox12.Text = (rdr["lname"].ToString());
                    guna2TextBox9.Text = (rdr["email"].ToString());
                    guna2TextBox11.Text = (rdr["number"].ToString());

                }
                tblSupplier.Close();
            }
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult1 = MessageBox.Show("Are you sure to cancel this???", "Cancel?", MessageBoxButtons.YesNo);
            if (dialogResult1 == DialogResult.Yes)
            {
                guna2TextBox2.Text = "";
                guna2TextBox3.Text = "";
                guna2TextBox1.Text = "";
                guna2TextBox4.Text = "";
                guna2TextBox5.Text = "";
                guna2TextBox6.Text = "";
                guna2TextBox8.Text = "";
                guna2TextBox7.Text = "";
                guna2TextBox10.Text = "";
                pictureBox1.Image = null;
                pictureBox2.Image = null;
                guna2TextBox13.Text = "";
                guna2TextBox12.Text = "";
                guna2TextBox9.Text = "";
                guna2TextBox11.Text = "";
            }
        }

        private void guna2TextBox15_TextChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox4.Checked == true)
            {
                (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format("EMPCODE LIKE '%{0}%'", guna2TextBox15.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

            }
            else if (guna2CheckBox3.Checked == true)
            {
                (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format("NAME LIKE '%{0}%'", guna2TextBox15.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
            }
            else
            {
                MessageBox.Show("Please select what do you want to search...");
            }
        }
    }
}
