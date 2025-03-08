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


namespace CSPSS.VOUCHER_MANAGE
{
    public partial class CLOSING_ACCOUNT : Form
    {
        DataTable dt = new DataTable();
        protected int i, j;
        protected string sql = @"SELECT * FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'";
        protected int M_int_judge, t;
        basec bc = new basec();
        XizheC.PERIOD period = new PERIOD();
        MAIN FRM1 = new MAIN ();
        private string _FINACIAL_YERAR;
        public string FINACIAL_YEAR
        {

            set { _FINACIAL_YERAR = value; }
            get { return _FINACIAL_YERAR; }

        }
        private string _PERIOD;
        public string PERIOD
        {
            set { _PERIOD = value; }
            get { return _PERIOD; }

        }
        private string _CYID;
        public string CYID
        {
            set { _CYID = value; }
            get { return _CYID; }

        }
        public CLOSING_ACCOUNT()
        {
            InitializeComponent();
           
        }
        public CLOSING_ACCOUNT(MAIN  FRM)
        {
            InitializeComponent();
            FRM1 = FRM;


        }
        private void FrmPeriodT_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            textBox1.ReadOnly = true;
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
        public  void bind()
        {
          
            dt = basec.getdts(sql);
            if (dt.Rows.Count > 0)
            {
               
                DateTime date1 = Convert.ToDateTime(dt.Rows[0]["FINANCIAL_YEAR"].ToString() + "/01/01");
                dateTimePicker1.Text = Convert.ToString(date1);
                DateTime date2 = Convert.ToDateTime("2014/"+dt.Rows[0]["PERIOD"].ToString() + "/01");
                dateTimePicker2.Text = Convert.ToString(date2);
                textBox1.Text = dt.Rows[0]["FINANCIAL_YEAR"].ToString() +"年第"+ dt.Rows[0]["PERIOD"].ToString()+"期";

                FINACIAL_YEAR = dt.Rows[0]["FINANCIAL_YEAR"].ToString();
                PERIOD = dt.Rows[0]["PERIOD"].ToString();
            }
            //textBox1.BackColor = Color.Yellow;
           
            FRM1.bind();
        }
        #endregion
        #region save
        protected void save()
        {
            btnSave.Focus();
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID;
            if (MessageBox.Show(@"确定要结账吗?结账后本期凭证不能再做修改删除操作！", "提示",
                                                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                if (period.GET_NUM_ID() == "")
                {

                }
                else
                {
                    string v1 = period.GET_NUM_ID();
                    SQlcommandE(period.sql);

                    basec.getcoms("UPDATE PERIOD SET IF_CURRENT_ACCOUNTING_PERIOD='N' WHERE PEID NOT IN ('" + v1 + "')");
                }
                bind();
            }
      


         
           /* CCURRENCY CCURRENCYo = new CCURRENCY();
            if (CCURRENCYo.NUMID == "")
            {

                return;
            }
            else if (CCURRENCYo.NUMID != "")
            {
              
                CYID = CCURRENCYo.NUMID;
            }
          
                CCURRENCY CCURRENCYx = new CCURRENCY();
                string CYKEY = CCURRENCYx.KEY;
                XizheC.PERIOD periodo = new PERIOD();
                string sql = @"INSERT INTO CCURRENCY_DET(CYKEY,CYID,PERIOD,INITIAL_RATE,CLOSING_RATE,YEAR,MONTH,DAY)
VALUES ('" + CYKEY + "','" + CYID + "','" +periodo .GETPERIOD + "','1.000','1.000','" + year +
    "','" + month + "','" + day + "')";
                basec.getcoms(sql);
            
            SQlcommandEo(CCURRENCYo.sql);*/
        }
        #region SQlcommandE
        protected void SQlcommandE(string sql)
        {
            XizheC.PERIOD periodo = new PERIOD();
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = "";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@PEID", SqlDbType.VarChar, 20).Value = period.GET_NUM_ID();

            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = periodo.NEXT_FINANCIAL_YEAR;
           
            sqlcom.Parameters.Add("@PERIOD", SqlDbType.VarChar, 20).Value = periodo.NEXT_PERIOD;

            sqlcom.Parameters.Add("@FINANCIAL_YEAR_INITIAL_DATE", SqlDbType.VarChar, 20).Value = "";
            sqlcom.Parameters.Add("@ACCOUNTING_PERIOD_START_DATE", SqlDbType.VarChar, 20).Value = periodo.ACCOUNTING_PERIOD_START_DATE;
            sqlcom.Parameters.Add("@ACCOUNTING_PERIOD_EXPIRATION_DATE", SqlDbType.VarChar, 20).Value = periodo.ACCOUNTING_PERIOD_EXPIRATION_DATE;
            sqlcom.Parameters.Add("@ACCOUNT_IF_START_USING", SqlDbType.VarChar, 20).Value = "";
            sqlcom.Parameters.Add("@IF_CURRENT_ACCOUNTING_PERIOD", SqlDbType.VarChar, 20).Value = "Y";
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion

        #region SQlcommandE
        protected void SQlcommandEo(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = "";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@CYID", SqlDbType.VarChar, 20).Value = CYID;
            sqlcom.Parameters.Add("@CYCODE", SqlDbType.VarChar, 20).Value = "RMB";
            sqlcom.Parameters.Add("@CYNAME", SqlDbType.VarChar, 20).Value = "人民币";
            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = period.FINANCIAL_YEAR;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
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
        private void btnSave_Click(object sender, EventArgs e)
        {
            save();
            try
            {
                
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
    
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
