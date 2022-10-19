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
    public partial class contacts : Form
    {


        private string cons = Properties.Settings.Default.dbcon; 
        public contacts()
        {
            InitializeComponent();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            using (SqlConnection tblOTS = new SqlConnection(cons))
            {
                tblOTS.Open();
                string insStmt2 = "insert into tblContacts ([name],[email]) values" +
                              " (@name,@email)";
                SqlCommand insCmd2 = new SqlCommand(insStmt2, tblOTS);
                insCmd2.Parameters.AddWithValue("@name", guna2TextBox1.Text);
                insCmd2.Parameters.AddWithValue("@email", guna2TextBox2.Text);
                insCmd2.ExecuteNonQuery();
                tblOTS.Close();
                obj.loaddgv();
                MessageBox.Show("Saved");
            }
        }
        Send_Email obj = (Send_Email)Application.OpenForms["Send_Email"];

        private void contacts_Load(object sender, EventArgs e)
        {

        }
    }
}
