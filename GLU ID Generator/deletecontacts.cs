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
    public partial class deletecontacts : Form
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
        public deletecontacts()
        {
            InitializeComponent();
        }

        private string cons = Properties.Settings.Default.dbcon;
        private void deletecontacts_Load(object sender, EventArgs e)
        {
            loaddgv();
        }
        Send_Email obj = (Send_Email)Application.OpenForms["Send_Email"];
        public void loaddgv()
        {
            try
            {
                using (SqlConnection codeMaterial = new SqlConnection(cons))
                {
                    dataGridView1.Rows.Clear();
                    DataTable dt = new DataTable();
                    codeMaterial.Open();
                    string list = "Select Id,name,email from tblContacts order by Id ASC";
                    SqlCommand command = new SqlCommand(list, codeMaterial);
                    SqlDataReader reader = command.ExecuteReader();
                    dt.Load(reader);
                    foreach (DataRow row in dt.Rows)
                    {
                        int a = dataGridView1.Rows.Add();
                        dataGridView1.Rows[a].Cells["Id"].Value = row["Id"].ToString();
                        dataGridView1.Rows[a].Cells["ename"].Value = row["name"].ToString();
                        dataGridView1.Rows[a].Cells["email"].Value = row["email"].ToString();
                    }

                    codeMaterial.Close();
                    dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
                    dataGridView1.RowHeadersVisible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                e.RowIndex >= 0)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure to delete this email?", "Delete?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    using (SqlConnection tblIn = new SqlConnection(cons))
                    {
                        tblIn.Open();
                        using (SqlCommand command = new SqlCommand("DELETE FROM tblContacts WHERE Id = '" + dataGridView1.CurrentRow.Cells["Id"].Value.ToString() + "'", tblIn))
                        {
                            command.ExecuteNonQuery();

                        }
                        tblIn.Close();
                        tblIn.Dispose();
                        MessageBox.Show("Deleted");
                        loaddgv();
                        obj.loaddgv();
                    }
                }
            }
        }
    }
}
