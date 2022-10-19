using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GLU_ID_Generator
{
    public partial class expiration : Form
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
        public expiration()
        {
            InitializeComponent();
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {

        }

        private void expiration_Load(object sender, EventArgs e)
        {

        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            saveToDatabase();
        }
        WorkersID obj = (WorkersID)Application.OpenForms["WorkersID"];
        private void saveToDatabase()
        {

            using (SqlConnection itemCode = new SqlConnection(con))
            {
                itemCode.Open();

                SqlCommand cmd = new SqlCommand("update tblSettings set ValidDate=@ValidDate,ExtendDate=@ExtendDate where Id=@Id", itemCode);

                cmd.Parameters.AddWithValue("@Id", "1");
                cmd.Parameters.AddWithValue("@ValidDate", guna2DateTimePicker2.Text);
                if (guna2CustomCheckBox1.Checked == true)
                {
                    cmd.Parameters.AddWithValue("@ExtendDate", guna2DateTimePicker3.Text);
                }
                else
                {
                    cmd.Parameters.AddWithValue("@ExtendDate", DBNull.Value);
                }
                cmd.ExecuteNonQuery();
                itemCode.Close();
                obj.dt();
                this.Close();
            }
        }
    }
}
