using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ronz.Pdf;
using Ronz.Core;
using Cyotek.GhostScript.PdfConversion;
using Cyotek.GhostScript;
using System.Web;
using Syncfusion.Pdf.Parsing;
using System.Reflection;
using System.Drawing.Drawing2D;

namespace GLU_ID_Generator
{
    public partial class editing : Form
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
        public string form;
        public editing()
        {
            InitializeComponent();
        }

        private void editing_Load(object sender, EventArgs e)
        {
            load();
            //dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //guna2ComboBox2.Text = "P.O Number";
            //this.dataGridView2.ColumnHeadersDefaultCellStyle.SelectionBackColor = this.dataGridView2.ColumnHeadersDefaultCellStyle.BackColor;

            //textBox2.BackColor = Color.FromArgb(169,169,169);

            ChangeControlStyles(dataGridView1, ControlStyles.OptimizedDoubleBuffer, true);
            this.dataGridView1.ColumnHeadersDefaultCellStyle.SelectionBackColor = this.dataGridView1.ColumnHeadersDefaultCellStyle.BackColor;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
        }
        private void ChangeControlStyles(Control ctrl, ControlStyles flag, bool value)
        {
            MethodInfo method = ctrl.GetType().GetMethod("SetStyle", BindingFlags.Instance | BindingFlags.NonPublic);
            if (method != null)
                method.Invoke(ctrl, new object[] { flag, value });
        }
        private void guna2TextBox1_DragOver(object sender, DragEventArgs e)
        {
            //if (e.Data.GetDataPresent(DataFormats.FileDrop))
            //    e.Effect = DragDropEffects.Link;
            //else
            //    e.Effect = DragDropEffects.None;
        }

        private void guna2TextBox1_DragDrop(object sender, DragEventArgs e)
        {
            //string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            //if (files != null && files.Any())
            //    guna2TextBox1.Text = files.First(); //select the first one 
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void guna2Panel1_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }
        string path;
        string extension;
        string name;

        private void guna2Panel1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            if (files != null && files.Any())
                path = files.First(); //select the first one
                                      //
            label1.Text = Path.GetFileName(path);
            //label2.Text = Path.GetExtension(path);
            extension = Path.GetExtension(path);
            name = Path.GetFileNameWithoutExtension(path);

            if (extension == ".pdf" || extension == ".PDF")
            {
                pictureBox1.Visible = true;
                pictureBox1.Image = Properties.Resources.icons8_pdf_480px_1;
                label1.Visible = true;
            }
            else
            {
                pictureBox1.Visible = true;
                pictureBox1.Image = Properties.Resources.icons8_error_144px_1;
                label1.Visible = true;
                label1.Text = "Please select PDF Only";
            }
        }

        private void guna2Panel1_MouseEnter(object sender, EventArgs e)
        {
        }

        private void guna2Panel1_MouseLeave(object sender, EventArgs e)
        {

        }

        private void guna2Panel1_DragEnter(object sender, DragEventArgs e)
        {
            pictureBox1.Visible = false;
            label1.Visible = false;
        }

        private void guna2Panel1_DragLeave(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            label1.Visible = true;
        }
        private void guna2Button6_Click(object sender, EventArgs e)
        {
            if (extension == ".pdf")
            {
                try
                {

                    using (PdfLoadedDocument loadedDocument = new PdfLoadedDocument(path))
                    {
                        for (int i = 0; i < loadedDocument.Pages.Count; i++)
                        {
                            //string n = name.Substring(0,5);
                            string paths = "../../ImgFolder/" + name;
                            DateTime date = DateTime.Now;
                            string yyyy = date.ToString("yyyy");
                            string MMM = date.ToString("MM");
                            string dd = date.ToString("dd");
                            string HH = date.ToString("HH");
                            string mm = date.ToString("mm");
                            string ss = date.ToString("ss");
                            paths += i.ToString() + MMM + dd + yyyy + HH + mm + ss;
                            Image img = loadedDocument.ExportAsImage(i);
                            img.Save(paths + ".png");
                            img.Dispose();
                            img = null;
                            paths = null;
                            //}

                            //else
                            //{


                            //}
                        }
                    }
                    //MessageBox.Show("All the pages are exported to output folder");
                    load();

                    MessageBox.Show("Save");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please select only PDF file");
            }
            
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

            load();
        }
        private void load()
        {
            dataGridView1.Columns.Clear();
            String[] files = System.IO.Directory.GetFiles(@"../../ImgFolder/", "*.png");
            //dataGridView1.Rows.Add();
            //DataTable table = new DataTable();
            //table.Columns.Add("Image");
            for (int i = 0; i < files.Length; i++)
            {
                //FileInfo file = new FileInfo(files[i]);
                Image img = new Bitmap(files[i]);
                //table.Rows.Add(img);
                DataGridViewImageColumn dgvimg = new DataGridViewImageColumn();
                Image dgvimage = Image.FromFile(files[i]);
                dgvimg.Image = dgvimage;
                dgvimg.ImageLayout = DataGridViewImageCellLayout.Zoom;
                dgvimg.Width = 200;
                dgvimg.HeaderText = Path.GetFileNameWithoutExtension(files[i]);
                dgvimg.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvimg.Name = Path.GetFileNameWithoutExtension(files[i]);
                dataGridView1.Columns.Add(dgvimg);
                if (dataGridView1.Rows.Count == 0)
                {
                    dataGridView1.Rows.Add();
                }
                //dataGridView1.Columns.Add(Path.GetFileNameWithoutExtension(files[i]), Path.GetFileNameWithoutExtension(files[i]));
                dataGridView1.Rows[0].Cells[Path.GetFileNameWithoutExtension(files[i])].Value = img;

            }

            //dataGridView1.DataSource = table;
        }

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
       
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            pictureBox3.Image = (Image)dataGridView1.CurrentCell.Value;
        }


        private void guna2Button2_Click(object sender, EventArgs e)
        {
            //radImageEditor1.ImageList.ImageStream = pictureBox1.Image;
        }

        private void pictureBox2_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void pictureBox2_MouseMove(object sender, MouseEventArgs e)
        {

        }
        WorkersID obj = (WorkersID)Application.OpenForms["WorkersID"];
        EmployeeForm obj1 = (EmployeeForm)Application.OpenForms["EmployeeForm"];

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            //WorkersID w = new WorkersID();
            if (form == "workers")
            {
                obj.pictureBox1.Image = pictureBox3.Image;
            }
            else
            {
                obj1.pictureBox1.Image = pictureBox3.Image;
            }
        }

        private void pictureBox2_MouseUp(object sender, MouseEventArgs e)
        {
        }

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                DataGridView.HitTestInfo info = dataGridView1.HitTest(e.X, e.Y);
                if (info.RowIndex >= 0)
                {
                    if (info.RowIndex >= 0 && info.ColumnIndex >= 0)
                    {
                        Image text = (Image)
                               dataGridView1.Rows[info.RowIndex].Cells[info.ColumnIndex].Value;
                        if (text != null)
                            dataGridView1.DoDragDrop(text, DragDropEffects.Copy);
                    }
                }
            }
        }

        private void editing_DragDrop(object sender, DragEventArgs e)
        {
          
        }

        private void radImageEditor1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(Bitmap)))
            {
                //radImageEditor1.OpenImage((Bitmap)e.Data.GetData(typeof(Bitmap)));
                //radImageEditor1.Refresh();
            }
        }

        private void radImageEditor1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void pictureBox2_DragDrop(object sender, DragEventArgs e)
        {
          
        }

        private void pictureBox2_DragEnter(object sender, DragEventArgs e)
        {
           
        }

        private void guna2Panel2_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            if (files != null && files.Any())
                path = files.First(); //select the first one
                                      //
                                      //label1.Text = Path.GetFileName(path);
                                      ////label2.Text = Path.GetExtension(path);
                                      //extension = Path.GetExtension(path);
                                      //name = Path.GetFileNameWithoutExtension(path);
                                      //FileInfo img = new FileInfo(path);
                                      //pictureBox3.Image = Bitmap.FromStream(img.Directory);
            if (path != null)
            {
                pictureBox3.Image = new Bitmap(path);
            }
            else
            {
                if (e.Data.GetDataPresent(typeof(Bitmap)))
                {
                    //radImageEditor1.OpenImage((Bitmap)e.Data.GetData(typeof(Bitmap)));
                    pictureBox3.Image = (Bitmap)e.Data.GetData(typeof(Bitmap));
                }
            }
            //if (extension == ".pdf" || extension == ".PDF")
            //{
            //    pictureBox1.Visible = true;
            //    pictureBox1.Image = Properties.Resources.icons8_pdf_480px_1;
            //    label1.Visible = true;
            //}
            //else
            //{
            //    pictureBox1.Visible = true;
            //    pictureBox1.Image = Properties.Resources.icons8_error_144px_1;
            //    label1.Visible = true;
            //    label1.Text = "Please select PDF Only";
            //}
        }

        private void guna2Panel2_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void guna2Button2_Click_1(object sender, EventArgs e)
        {
            string path = @"..\..\ImgFolder\";
            System.Diagnostics.Process.Start(path);
        }

        private void radImageEditor1_ImageSaved(object sender, EventArgs e)
        {
            load();
        }
    }
}
