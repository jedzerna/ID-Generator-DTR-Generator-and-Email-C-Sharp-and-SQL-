using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class WorkersID : Form
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
        public WorkersID()
        {
            InitializeComponent();
        }

        private void WorkersID_Load(object sender, EventArgs e)
        {
            load();
            loadhist();
            guna2DateTimePicker1.Value = DateTime.Now;
            dt();
        }
       
        public void dt()
        {
            using (SqlConnection tblSupplier = new SqlConnection(con))
            {
                tblSupplier.Open();
                DataTable dt = new DataTable();
                String query = "SELECT * FROM tblSettings WHERE Id = 1";
                SqlCommand cmd = new SqlCommand(query, tblSupplier);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    label20.Text = (rdr["ValidDate"].ToString());
                    label19.Text = (rdr["ExtendDate"].ToString());
                }
                tblSupplier.Close();
            }
        }
        string lb = "{";
        string rb = "}";
        public void load()
        {
            using (SqlConnection tblOTS = new SqlConnection(con))
            {
                DataTable employeename = new DataTable();
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

                        //int c = dataGridView1.Rows.Add();
                        string first = string.Concat(item["FIRST"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
                        string middle = string.Concat(item["MIDDLE"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
                        string family = string.Concat(item["FAMILY"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));



                        drLocal = employeename.NewRow();
                        drLocal["EMPCODE"] = item["EMPCODE"].ToString();
                        drLocal["NAME"] = family + ", " + first + " " + middle;
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
                    string list = "SELECT Id,EMPCODE,FNAME,MNAME,LNAME,DATE FROM tblWorkersHist";
                    SqlDataAdapter command = new SqlDataAdapter(list, dbDR);
                    command.Fill(dt);
                    foreach (DataRow item in dt.Rows)
                    {

                        //int c = dataGridView1.Rows.Add();
                        string first = string.Concat(item["FNAME"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "'"));
                        string middle = string.Concat(item["MNAME"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
                        string family = string.Concat(item["LNAME"].ToString().Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

                        drLocal = employeename.NewRow();
                    drLocal["Id"] = item["Id"].ToString();
                    drLocal["EMPCODE"] = item["EMPCODE"].ToString();
                        drLocal["NAME"] = family + ", " + first + " " + middle;
                    drLocal["DATE"] = item["DATE"].ToString();
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

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {

            this.WindowState = FormWindowState.Minimized;
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            saveToDatabase();

            loadhist();
        }
        private void saveToDatabase()
        {
            if (pictureBox1.Image == null)
            {
                MessageBox.Show("Please add an Image of the ID..");
            }
            else
            {
                using (SqlConnection itemCode = new SqlConnection(con))
                {

                    MemoryStream ms = new MemoryStream();
                    pictureBox1.Image.Save(ms, pictureBox1.Image.RawFormat);
                    byte[] pic1 = ms.ToArray();


                    itemCode.Open();

                    SqlCommand cmd = new SqlCommand("update tblWorkers set NAME=@NAME,PROJECT=@PROJECT,EMPCODE=@EMPCODE,DESIGNATION=@DESIGNATION,ENAME=@ENAME,ESTREET=@ESTREET,ECITY=@ECITY,ENUMBER=@ENUMBER,BDAY=@BDAY,SSS=@SSS,VALIDUNTIL=@VALIDUNTIL,EXTEND=@EXTEND,IMAGE=@IMAGE where Id=@Id", itemCode);

                    cmd.Parameters.AddWithValue("@Id", "1");
                    cmd.Parameters.AddWithValue("@NAME", guna2TextBox12.Text + ", " + guna2TextBox2.Text + " " + guna2TextBox13.Text);
                    cmd.Parameters.AddWithValue("@PROJECT", guna2TextBox3.Text);
                    cmd.Parameters.AddWithValue("@EMPCODE", guna2TextBox1.Text);
                    cmd.Parameters.AddWithValue("@DESIGNATION", guna2ComboBox1.Text);
                    cmd.Parameters.AddWithValue("@ENAME", guna2TextBox4.Text);
                    cmd.Parameters.AddWithValue("@ESTREET", guna2TextBox6.Text);
                    cmd.Parameters.AddWithValue("@ECITY", guna2TextBox8.Text);
                    cmd.Parameters.AddWithValue("@ENUMBER", guna2TextBox5.Text);
                    cmd.Parameters.AddWithValue("@BDAY", guna2DateTimePicker1.Text);
                    cmd.Parameters.AddWithValue("@SSS", guna2TextBox7.Text);
                    cmd.Parameters.AddWithValue("@VALIDUNTIL", label20.Text);
                    cmd.Parameters.AddWithValue("@EXTEND", label19.Text);
                    cmd.Parameters.AddWithValue("@IMAGE", SqlDbType.Image).Value = pic1;
                    cmd.ExecuteNonQuery();

                    itemCode.Close();

                    itemCode.Open();
                    string insStmt = "insert into tblWorkersHist ([FNAME],[PROJECT],[EMPCODE],[DESIGNATION],[ENAME],[ESTREET],[ECITY],[ENUMBER],[BDAY],[SSS],[VALIDUNTIL],[EXTEND],[IMAGE],[MNAME],[LNAME],[DATE])" +
                        " values (@FNAME,@PROJECT,@EMPCODE,@DESIGNATION,@ENAME,@ESTREET,@ECITY,@ENUMBER,@BDAY,@SSS,@VALIDUNTIL,@EXTEND,@IMAGE,@MNAME,@LNAME,@DATE)";
                    SqlCommand insCmd = new SqlCommand(insStmt, itemCode);
                    insCmd.Parameters.AddWithValue("@FNAME", guna2TextBox2.Text);
                    insCmd.Parameters.AddWithValue("@PROJECT", guna2TextBox3.Text);
                    insCmd.Parameters.AddWithValue("@EMPCODE", guna2TextBox1.Text);
                    insCmd.Parameters.AddWithValue("@DESIGNATION", guna2ComboBox1.Text);
                    insCmd.Parameters.AddWithValue("@ENAME", guna2TextBox4.Text);
                    insCmd.Parameters.AddWithValue("@ESTREET", guna2TextBox6.Text);
                    insCmd.Parameters.AddWithValue("@ECITY", guna2TextBox8.Text);
                    insCmd.Parameters.AddWithValue("@ENUMBER", guna2TextBox5.Text);
                    insCmd.Parameters.AddWithValue("@BDAY", guna2DateTimePicker1.Text);
                    insCmd.Parameters.AddWithValue("@SSS", guna2TextBox7.Text);
                    insCmd.Parameters.AddWithValue("@VALIDUNTIL", label20.Text);
                    insCmd.Parameters.AddWithValue("@EXTEND", label19.Text);
                    insCmd.Parameters.AddWithValue("@IMAGE", SqlDbType.Image).Value = pic1;
                    insCmd.Parameters.AddWithValue("@MNAME", guna2TextBox13.Text);
                    insCmd.Parameters.AddWithValue("@LNAME", guna2TextBox12.Text);

                    DateTime dat = DateTime.Now;
                    string fdate = dat.ToString("MM/dd/yyyy HH:mm");
                    insCmd.Parameters.AddWithValue("@DATE", fdate);
                    int affectedRows = insCmd.ExecuteNonQuery();

                    itemCode.Close();

                    Form1 f = new Form1();
                    f.idtype = "worker";
                    f.ShowDialog();
                }
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

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            expiration ee = new expiration();
            ee.ShowDialog();
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
                            load();
                        }
                    }
                    else
                    {
                        MessageBox.Show("The PERID.DBF doesn't exist make sure to locate the file with the specified folder.");
                    }
                }
            }
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
          
        }

        private void guna2TextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {
        }

        private void guna2TextBox14_TextChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox1.Checked == true)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("EMPCODE LIKE '%{0}%'", guna2TextBox14.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

            }
            else if (guna2CheckBox2.Checked == true)
            {
                (dataGridView1.DataSource as DataTable).DefaultView.RowFilter = string.Format("NAME LIKE '%{0}%'", guna2TextBox14.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
            }
            else
            {
                MessageBox.Show("Please select what do you want to search...");
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

        private void WorkersID_MouseDown(object sender, MouseEventArgs e)
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
                    //guna2TextBox8.Text = (rdr["BDAY"].ToString());
                    if (rdr["BDAY"].ToString() != "" || rdr["BDAY"] != DBNull.Value)
                    {
                        DateTime bday = DateTime.Parse(rdr["BDAY"].ToString());
                        guna2DateTimePicker1.Value = bday;
                    }
                    //DateTime date = bday("MMMM dd, yyyy");
                    
                }
                tblSupplier.Close();
            }
        }

        private void WorkersID_FormClosed(object sender, FormClosedEventArgs e)
        {
            Selection d=new Selection();
            this.Hide();
            d.ShowDialog();
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            using (SqlConnection tblSupplier = new SqlConnection(con))
            {
                tblSupplier.Open();
                DataTable dt = new DataTable();
                String query = "SELECT * FROM tblWorkersHist WHERE Id = '" + dataGridView2.CurrentRow.Cells["Id"].Value.ToString() + "'";
                SqlCommand cmd = new SqlCommand(query, tblSupplier);
                SqlDataReader rdr = cmd.ExecuteReader();

                if (rdr.Read())
                {
                    guna2TextBox2.Text = (rdr["FNAME"].ToString());
                    guna2TextBox3.Text = (rdr["PROJECT"].ToString());
                    guna2TextBox1.Text = (rdr["EMPCODE"].ToString());
                    guna2ComboBox1.Text = (rdr["DESIGNATION"].ToString());
                    guna2TextBox4.Text = (rdr["ENAME"].ToString());
                    guna2TextBox6.Text = (rdr["ESTREET"].ToString());
                    guna2TextBox8.Text = (rdr["ECITY"].ToString());
                    guna2TextBox5.Text = (rdr["ENUMBER"].ToString());
                    guna2DateTimePicker1.Text = (rdr["BDAY"].ToString());
                    guna2TextBox7.Text = (rdr["SSS"].ToString());
                    label20.Text = (rdr["VALIDUNTIL"].ToString());

                    if (rdr["IMAGE"] != DBNull.Value)
                    {
                        byte[] img = (byte[])(rdr["IMAGE"]);
                        MemoryStream mstream = new MemoryStream(img);
                        pictureBox1.Image = System.Drawing.Image.FromStream(mstream);
                    }
                    //guna2TextBox13.Text = (rdr["IMAGE"].ToString());
                    label19.Text = (rdr["EXTEND"].ToString());
                    guna2TextBox13.Text = (rdr["MNAME"].ToString());
                    guna2TextBox12.Text = (rdr["LNAME"].ToString());


                }
                tblSupplier.Close();
            }
        }

        private void guna2Button5_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult1 = MessageBox.Show("Are you sure to cancel this???", "Cancel?", MessageBoxButtons.YesNo);
            if (dialogResult1 == DialogResult.Yes)
            {
                guna2TextBox2.Text = "";
                guna2TextBox3.Text = "";
                guna2TextBox1.Text = "";
                guna2ComboBox1.Text = "";
                guna2TextBox4.Text = "";
                guna2TextBox6.Text = "";
                guna2TextBox8.Text = "";
                guna2TextBox5.Text = "";
                guna2DateTimePicker1.Text = "";
                guna2TextBox7.Text = "";
                label20.Text = "";

                pictureBox1.Image = null;

                //guna2TextBox13.Text = (rdr["IMAGE"].ToString());
                label19.Text = "";
                guna2TextBox13.Text = "";
                guna2TextBox12.Text = "";
                load();
                loadhist();
                guna2DateTimePicker1.Value = DateTime.Now;
                dt();
            }
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

        private void guna2TextBox9_TextChanged(object sender, EventArgs e)
        {
            if (guna2CheckBox4.Checked == true)
            {
                (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format("EMPCODE LIKE '%{0}%'", guna2TextBox9.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));

            }
            else if (guna2CheckBox3.Checked == true)
            {
                (dataGridView2.DataSource as DataTable).DefaultView.RowFilter = string.Format("NAME LIKE '%{0}%'", guna2TextBox9.Text.Replace("[", lb).Replace("]", rb).Replace("​*", "[*​]").Replace("%", "[%]").Replace("'", "''"));
            }
            else
            {
                MessageBox.Show("Please select what do you want to search...");
            }
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
            editing b = new editing();
            b.form = "workers";
            b.ShowDialog();
        }
    }
}
