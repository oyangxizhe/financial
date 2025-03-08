using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using XizheC;


namespace CSPSS.BASE_INFO
{
    public partial class PAYMENT_METHOD  : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        basec boperate = new basec();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        protected int M_int_judge, t;
        basec bc = new basec();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        CPAYMENT_METHOD cpayment_method = new CPAYMENT_METHOD();
        string sql = @"
SELECT A.PMID AS 代码,
A.PMNAME AS 结算方式,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID ) AS 制单人,
A.DATE AS  制单日期 
FROM PAYMENT_METHOD A
";
        string sql1 = @"INSERT INTO PAYMENT_METHOD(
PMID,
PMNAME,
MAKERID,
DATE,
YEAR,
MONTH
) VALUES 

(
@PMID,
@PMNAME,
@MAKERID,
@DATE,
@YEAR,
@MONTH
)

";
        string sql2 = @"UPDATE PAYMENT_METHOD SET 
PMID=@PMID,
PMNAME=@PMNAME,
MAKERID=@MAKERID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH
";
        public PAYMENT_METHOD()
        {
            InitializeComponent();
        }

        private void PAYMENT_METHOD_Load(object sender, EventArgs e)
        {
            textBox2.BorderStyle = BorderStyle.FixedSingle;
            try
            {

                bind();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
   
        #region bind
        private void bind()
        {

            textBox1.Text = IDO;
            dt = basec.getdts(sql);
            dataGridView1.DataSource = dt;
            textBox2.Focus();
            textBox2.BackColor = Color.Yellow;
            dgvStateControl();
            hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
               
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }

            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 6)
                {
                    dataGridView1.Columns[i].ReadOnly = true;
                }
                else
                {
                    dataGridView1.Columns[i].ReadOnly = true;
                }
                if (i == 0)
                {
                    dataGridView1.Columns[i].Visible = true;

                }
            }
        }
        #endregion
       
        #region SQlcommandE
        protected void SQlcommandE(string sql, string IDVALUE, string NAMEVALUE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = "";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@PMID", SqlDbType.VarChar, 20).Value = IDVALUE;
            sqlcom.Parameters.Add("@PMNAME", SqlDbType.VarChar, 20).Value = NAMEVALUE;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&
             (
             (
              !(ActiveControl is System.Windows.Forms.TextBox) ||
              !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)
             )
             )
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
      
  
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.0000";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }

            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("数量只能输入数字！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region excelprint
        private void btnExcelPrint_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dtn = boperate.PrintOrder(" WHERE ORID='" + textBox1.Text + "'");
                if (dtn.Rows.Count > 0)
                {
                    string v1 = @"D:\PrintModelForOrder.xls";
                    if (File.Exists(v1))
                    {
                        boperate.ExcelPrint(dtn, "订单", v1);
                    }
                    else
                    {
                        MessageBox.Show("指定路径不存在打印模版！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                else
                {
                    MessageBox.Show("无数据可打印！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        #region btnadd
        private void btnAdd_Click(object sender, EventArgs e)
        {
            ClearText();
            hint.Text = "";
            textBox1.Text = cpayment_method.GETID();
           
        }
        #endregion
        private void ClearText()
        {
            textBox2.Text = "";
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
           
            try
            {
                if (juage())
                {

                }
                else
                {
                    save("PAYMENT_METHOD", "PMID", "PMNAME", textBox1.Text, textBox2.Text, "代码", "结算方式");
                    if (IFExecution_SUCCESS)
                    {
                        bind();
                        add();
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region juage()
        private bool juage()
        {
         
            bool b = false;
            if (textBox1.Text == "")
            {
                b = true;
                hint.Text = "编号不能为空！";
            }
            else if (textBox2.Text == "")
            {
                b = true;
                hint.Text = "结算方式不能为空！";
            }
            return b;

        }
        #endregion
        private void add()
        {

            textBox1.Text = cpayment_method.GETID();
            ClearText();
            textBox2.Focus();

        }
         public void save(string TABLENAME,string COLUMNID,string COLUMNNAME,string IDVALUE,string NAMEVALUE,string INFOID,string INFONAME)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE "+COLUMNID +"='"+IDVALUE+"'" );
            string v2 = bc.getOnlyString("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE "+COLUMNNAME+"='"+NAMEVALUE +"'");
            //string varMakerID;
            if (!bc.exists("SELECT "+COLUMNID+" FROM "+TABLENAME+" WHERE "+COLUMNID +"='"+IDVALUE+"'" ))
            {
                if (bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE "+COLUMNNAME +"='"+NAMEVALUE+"'"))
                {

                    hint.Text = INFONAME + "已经存在于系统！";
                    IFExecution_SUCCESS = false;

                }
                else
                {

                    SQlcommandE(sql1,IDVALUE ,NAMEVALUE);
                    IFExecution_SUCCESS = true;

                }

            }
 
            else if (v2 !=NAMEVALUE)
            {
                if (bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE " + COLUMNNAME + "='" + NAMEVALUE + "'"))
                {

                    hint.Text = INFONAME + "已经存在于系统！";
                    IFExecution_SUCCESS = false;

                }
                else
                {

                    SQlcommandE(sql2+" WHERE "+COLUMNID+"='"+IDVALUE+"'" ,IDVALUE ,NAMEVALUE );
                    IFExecution_SUCCESS = true;

                }
            }
  
            else
            {

                SQlcommandE(sql2 + " WHERE "+COLUMNID +"='" + IDVALUE+"'", IDVALUE, NAMEVALUE);
                IFExecution_SUCCESS = true;

            }
         
        }
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #region btndel
        private void btnDel_Click(object sender, EventArgs e)
        {
           
     
            try
            {

                if (MessageBox.Show("确定要删除该条信息吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {

                    basec.getcoms("DELETE PAYMENT_METHOD WHERE PMID='" + Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim() + "'");
                    bind();
                    //ClearText();
                    //textBox1.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        #region dgvcellclick
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
          

            try
            {

                textBox1.Text = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                textBox2.Text = Convert.ToString(dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value).Trim();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
        #endregion
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                dt = Search();
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dgvStateControl();
                }
                else
                {
                    MessageBox.Show("没有要查找的相关记录！");
                    dataGridView1.DataSource = null;
                }

                dgvStateControl();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region Search()
        private  DataTable Search()
        {

            string sql1 = @" where A.PMNAME like '%" + textBox2.Text + "%'";
            dt = basec.getdts(sql + sql1);
            return dt;
        }
        #endregion
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, "结算方式信息");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
