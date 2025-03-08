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
    public partial class THE_FINAL_BILL : Form
    {
        DataTable dt = new DataTable();
        protected int i, j;
        protected string sql = @"SELECT * FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'";
        protected int M_int_judge, t;
        CTHE_FINAL_BILL the_final_bill = new CTHE_FINAL_BILL();
        basec bc = new basec();
        XizheC.PERIOD period = new PERIOD();
        MAIN FRM1 = new MAIN();
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
        public THE_FINAL_BILL()
        {
            InitializeComponent();
           
        }
        public THE_FINAL_BILL(MAIN FRM)
        {
            InitializeComponent();
            FRM1 = FRM;


        }
        private void PeriodT_Load(object sender, EventArgs e)
        {
           
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
            if (MessageBox.Show(@"确定要结转吗?结转后本期凭证不能再做修改删除操作！", "提示",
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
      


         
           /* CURRENCY currencyo = new CURRENCY();
            if (currencyo.NUMID == "")
            {

                return;
            }
            else if (currencyo.NUMID != "")
            {
              
                CYID = currencyo.NUMID;
            }
          
                CURRENCY currencyx = new CURRENCY();
                string CYKEY = currencyx.KEY;
                XizheC.PERIOD periodo = new PERIOD();
                string sql = @"INSERT INTO CURRENCY_DET(CYKEY,CYID,PERIOD,INITIAL_RATE,CLOSING_RATE,YEAR,MONTH,DAY)
VALUES ('" + CYKEY + "','" + CYID + "','" +periodo .GETPERIOD + "','1.000','1.000','" + year +
    "','" + month + "','" + day + "')";
                basec.getcoms(sql);
            
            SQlcommandEo(currencyo.sql);*/
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
            PERIOD peroid = new PERIOD();
            DataTable dtx = new DataTable();
            if (peroid.IF_CARRY)
            {
                MessageBox.Show("当期已经结转过损益！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (!bc.exists("SELECT * FROM VOUCHER_MST WHERE FINANCIAL_YEAR='" + peroid.FINANCIAL_YEAR + "' AND PERIOD='" + peroid.GETPERIOD + "'"))
            {
                MessageBox.Show("当期无数据可结转！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {


                the_final_bill.MAKER_VOUCHER(@" 
WHERE  B.FINANCIAL_YEAR='" + peroid.FINANCIAL_YEAR + "' AND B.PERIOD='" + peroid.GETPERIOD + "'  AND B.STATUS IN ('OPEN','INITIAL','CARRY')  ORDER BY C.ACCODE ASC  ");
            }
        }
    
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
