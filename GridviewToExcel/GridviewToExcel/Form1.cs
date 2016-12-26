using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Common;

using System.Data.SqlClient;
using System.Configuration;
using Microsoft.Office.Interop.Excel;





namespace GridviewToExcel
{
    public partial class Form1 : Form
    {
        public string MyConnectionStrings = ConfigurationManager.ConnectionStrings["INC_SVH_connection"].ConnectionString;
      

        public Form1()
        {
            InitializeComponent();
        }

        private void LoadGrid_Click(object sender, EventArgs e)
        {

            DataSet ds = new DataSet();
            // System.Data.DataTable dt = new System.Data.DataTable();
            System.Data.DataTable dt = ds.Tables.Add("Orders");
            
            SqlConnection con = new SqlConnection(MyConnectionStrings);
           

            string sql_inc = @"SELECT [A_DOC_NO],[A_DTE],[A_DTE],[A_DTE_FRM],[A_TME],[A_INC_CAT],[A_INC_CAT_DES],[A_DEP],[A_DEP_DES],[A_INC_TYP],[A_INC_TYP_DES]
,[A_INC_SUB],[A_INC_SUB_DES],[A_DEP2_DES1],[A_DEP2_DES2],[A_DEP2_DES3],[A_DEP2_DES4],[A_DEP2_DES5],[A_Committee1],[A_Committee2],[A_Committee3]
,[A_RN],[A_RMK],[A_SVR],[A_SVR_DES],[A_SLV],[A_INC_STS_DES],[A_NTE1],[A_NTE2],[A_NTE3],[A_NTE1],[A_NTE2],[A_NTE3]
                            FROM INCIDENTANN";
            SqlDataAdapter da = new SqlDataAdapter(sql_inc, con);
            da.Fill(ds, "IncedentReport");

            
            //set column in datatable 
            dt.Columns.Add("Doc_number", typeof(string));
            dt.Columns.Add("Month", typeof(string));
            dt.Columns.Add("Date_key", typeof(DateTime));
            dt.Columns.Add("Date_Occur", typeof(DateTime));
            dt.Columns.Add("Time_Occur", typeof(DateTime));

            dt.Columns.Add("CategoryCode", typeof(string));
            dt.Columns.Add("CategoryDesc", typeof(string));
            dt.Columns.Add("Dept_key", typeof(string));
            dt.Columns.Add("Dept_key_Desc", typeof(string));
            dt.Columns.Add("Type", typeof(string));

            dt.Columns.Add("Type_Desc", typeof(string));
            dt.Columns.Add("SubType", typeof(string));
            dt.Columns.Add("SubType_Desc", typeof(string));
            dt.Columns.Add("Dept_Concern1", typeof(string));
            dt.Columns.Add("Dept_Concern2", typeof(string));

            dt.Columns.Add("Dept_Concern3", typeof(string));
            dt.Columns.Add("Dept_Concern4", typeof(string));
            dt.Columns.Add("Dept_Concern5", typeof(string));
            dt.Columns.Add("Committee1", typeof(string));
            dt.Columns.Add("Committee2", typeof(string));

            dt.Columns.Add("Committee3", typeof(string));
            dt.Columns.Add("HN", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            dt.Columns.Add("Severity", typeof(string));
            dt.Columns.Add("SeverityDesc", typeof(string));

            dt.Columns.Add("Solve_Primary", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("Response1", typeof(string));
            dt.Columns.Add("Response2", typeof(string));
            dt.Columns.Add("Response3", typeof(string));

            dt.Columns.Add("CommentByRM", typeof(string));
            dt.Columns.Add("RCA_Code", typeof(string));
            dt.Columns.Add("RCA_Desc", typeof(string));

            

            //dataGridView1.DataSource = dt;
            dataGridView1.DataSource = ds.Tables["IncedentReport"];


            label3.Text = datefrom.Value.ToString("dd-MMM-yyyy");
            label4.Text = dateto.Value.ToString("dd-MMM-yyyy");

        }

        private void Ex2Excel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = Excel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)Excel.ActiveSheet;
            ws.Rows.RowHeight = 15;
            Excel.Visible = true;
   
            ws.Cells[1, 1] = "Doc_number";
            ws.Cells[1, 2] = "Month";
            ws.Cells[1, 3] = "Date_key";
            ws.Cells[1, 4] = "Date_Occur";
            ws.Cells[1, 5] = "Time_Occur";
            ws.Cells[1, 6] = "CategoryCode";
            ws.Cells[1, 7] = "CategoryDesc";
            ws.Cells[1, 8] = "Dept_key";
            ws.Cells[1, 9] = "Dept_key_Desc";
            ws.Cells[1, 10] = "Type";
            ws.Cells[1, 11] = "Type_Desc";
            ws.Cells[1, 12] = "SubType";
            ws.Cells[1, 13] = "SubType_Desc";
            ws.Cells[1, 14] = "Dept_Concern1";
            ws.Cells[1, 15] = "Dept_Concern2";
            ws.Cells[1, 16] = "Dept_Concern3";
            ws.Cells[1, 17] = "Dept_Concern4";
            ws.Cells[1, 18] = "Dept_Concern5";
            ws.Cells[1, 19] = "Committee1";
            ws.Cells[1, 20] = "Committee2";
            ws.Cells[1, 21] = "Committee3";
            ws.Cells[1, 22] = "HN";
            ws.Cells[1, 23] = "Remark";
            ws.Cells[1, 24] = "Severity";
            ws.Cells[1, 25] = "SeverityDesc";
            ws.Cells[1, 26] = "Solve_Primary";
            ws.Cells[1, 27] = "Status";
            ws.Cells[1, 28] = "Response1";
            ws.Cells[1, 29] = "Response2";
            ws.Cells[1, 30] = "Response3";
            ws.Cells[1, 31] = "CommentByRM";
            ws.Cells[1, 32] = "RCA_Code";
            ws.Cells[1, 33] = "RCA_Desc";
            

            for (int j = 2; j <= dataGridView1.Rows.Count; j++)
            {
                for (int i = 1; i <= 33; i++)
                {
                    ws.Cells[j,i] =dataGridView1.Rows[j-2].Cells[i-1].Value;
                }
            }

        }
    }
}
