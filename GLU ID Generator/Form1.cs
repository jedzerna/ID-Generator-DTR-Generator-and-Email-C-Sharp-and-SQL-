using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace GLU_ID_Generator
{
    public partial class Form1 : Form
    {
        public string idtype;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (idtype == "worker")
            {
                WindowState = FormWindowState.Maximized;
                reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
                var pagesetting = new System.Drawing.Printing.PageSettings();
                pagesetting.Margins = new System.Drawing.Printing.Margins(20, 20, 20, 20);
                //pagesetting.Landscape = new System.Drawing.Printing.lan
                string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\workersID.rdlc";

                ReportDataSource dr = new ReportDataSource("DataSet1", WORKERSDATA());
                this.reportViewer1.LocalReport.ReportPath = exeFolder;
                this.reportViewer1.LocalReport.DataSources.Add(dr);
                this.reportViewer1.SetPageSettings(pagesetting);
                this.reportViewer1.RefreshReport();
            }
            else if (idtype == "GLU")
            {
                WindowState = FormWindowState.Maximized;
                reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
                var pagesetting = new System.Drawing.Printing.PageSettings();
                pagesetting.Margins = new System.Drawing.Printing.Margins(50, 0, 50, 0);

                string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\PrintID.rdlc";

                ReportDataSource dr = new ReportDataSource("EMPINFO", projectsumdataa());
                this.reportViewer1.LocalReport.ReportPath = exeFolder;
                this.reportViewer1.LocalReport.DataSources.Add(dr);
                this.reportViewer1.SetPageSettings(pagesetting);
                this.reportViewer1.RefreshReport();
            }
            else if(idtype == "glubros")
            {
                WindowState = FormWindowState.Maximized;
                reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
                var pagesetting = new System.Drawing.Printing.PageSettings();
                pagesetting.Margins = new System.Drawing.Printing.Margins(50, 0, 50, 0);

                string exeFolder = Path.GetDirectoryName(Application.ExecutablePath) + @"\Report1.rdlc";

                ReportDataSource dr = new ReportDataSource("EMPINFO", projectsumdataa());
                this.reportViewer1.LocalReport.ReportPath = exeFolder;
                this.reportViewer1.LocalReport.DataSources.Add(dr);
                this.reportViewer1.SetPageSettings(pagesetting);
                this.reportViewer1.RefreshReport();
            }

        }
        private string con = Properties.Settings.Default.dbcon;
        private DataTable projectsumdataa()
        {
            using (SqlConnection dbDR = new SqlConnection(con))
            {
                DataTable dt = new DataTable();
                dbDR.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM tblEmp", dbDR);
                SqlDataReader rd = cmd.ExecuteReader();
                dt.Load(rd);
                dbDR.Close();
                return dt;
            }

        }
        private DataTable WORKERSDATA()
        {
            using (SqlConnection dbDR = new SqlConnection(con))
            {
                DataTable dt = new DataTable();
                dbDR.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM tblWorkers", dbDR);
                SqlDataReader rd = cmd.ExecuteReader();
                dt.Load(rd);
                dbDR.Close();
                return dt;
            }

        }
    }
}
