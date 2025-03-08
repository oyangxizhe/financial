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
    public partial class INITIAL_PERIOD : Form
    {
        DataTable dt = new DataTable();
        protected int i, j;
        protected string sql = @"SELECT * FROM PERIOD";
        protected int M_int_judge, t;
        basec bc = new basec();
        XizheC.PERIOD period = new PERIOD();
        private string _FINANCIAL_YEAR_INITIAL_DATE;
        public string FINANCIAL_YEAR_INITIAL_DATE
        {

            set { _FINANCIAL_YEAR_INITIAL_DATE = value; }
            get { return _FINANCIAL_YEAR_INITIAL_DATE; }

        }
        private string _ACCOUNTING_PERIOD_START_DATE;
        public string ACCOUNTING_PERIOD_START_DATE
        {
            set { _ACCOUNTING_PERIOD_START_DATE = value; }
            get { return _ACCOUNTING_PERIOD_START_DATE; }
        }
        private string _ACCOUNTING_PERIOD_EXPIRATION_DATE;
        public string ACCOUNTING_PERIOD_EXPIRATION_DATE
        {

            set { _ACCOUNTING_PERIOD_EXPIRATION_DATE = value; }
            get { return _ACCOUNTING_PERIOD_EXPIRATION_DATE; }

        }
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

        string sql2 = @"UPDATE Period SET 
PEID=@PEID,
FINANCIAL_YEAR_INITIAL_DATE=@FINANCIAL_YEAR_INITIAL_DATE,
ACCOUNTING_PERIOD_START_DATE=@ACCOUNTING_PERIOD_START_DATE,
FINANCIAL_YEAR=@FINANCIAL_YEAR,
PERIOD=@PERIOD,
ACCOUNT_IF_START_USING=@ACCOUNT_IF_START_USING,
MAKERID=@MAKERID,
DATE=@DATE,
IF_CURRENT_ACCOUNTING_PERIOD=@IF_CURRENT_ACCOUNTING_PERIOD,
ACCOUNTING_PERIOD_EXPIRATION_DATE=@ACCOUNTING_PERIOD_EXPIRATION_DATE
";
        string sql3 = @"
INSERT INTO CURRENCY_DET(
CYKEY,
CYID,
PERIOD,
INITIAL_RATE,
CLOSING_RATE,
YEAR,
MONTH,
DAY
)
VALUES
(
@CYKEY,
@CYID,
@PERIOD,
@INITIAL_RATE,
@CLOSING_RATE,
@YEAR,
@MONTH,
@DAY

)

";
        string sql4 = @"
INSERT INTO CURRENCY_MST(
CYID,
CYCODE,
CYNAME,
FINANCIAL_YEAR,
MAKERID,
DATE,
YEAR,
MONTH

)
VALUES
(
@CYID,
@CYCODE,
@CYNAME,
@FINANCIAL_YEAR,
@MAKERID,
@DATE,
@YEAR,
@MONTH


)
";

        MAIN F1 = new MAIN();
        public INITIAL_PERIOD()
        {
            InitializeComponent();

        }
        public INITIAL_PERIOD(MAIN  FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void INITIAL_PERIOD_Load(object sender, EventArgs e)
        {

            bind();
            try
            {
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #region bind
        public void bind()
        {
            numericUpDown1.Value = 1;
            numericUpDown1.Maximum = 12;
            numericUpDown1.Minimum = 1;
            numericUpDown1.TextAlign = HorizontalAlignment.Center;

            numericUpDown2.Value = 1;
            numericUpDown2.Maximum = 28;
            numericUpDown2.Minimum = 1;
            numericUpDown2.TextAlign = HorizontalAlignment.Center;

            numericUpDown3.Maximum = 5000;
            numericUpDown3.Value = Convert.ToInt32(DateTime.Now.Date.ToString("yyyy"));

            numericUpDown3.Minimum = 1900;
            numericUpDown3.TextAlign = HorizontalAlignment.Center;

            numericUpDown4.Value = Convert.ToInt32(DateTime.Now.Date.ToString("MM"));
            numericUpDown4.Maximum = 12;
            numericUpDown4.Minimum = 1;
            numericUpDown4.TextAlign = HorizontalAlignment.Center;

            //numericUpDown1.ReadOnly = true;
            //numericUpDown1.BackColor = Color.White;
            dt = basec.getdts(sql);
            if (dt.Rows.Count > 0)
            {
                string v = dt.Rows[0]["FINANCIAL_YEAR_INITIAL_DATE"].ToString();
                FINACIAL_YEAR = dt.Rows[0]["FINANCIAL_YEAR"].ToString();
                PERIOD = dt.Rows[0]["PERIOD"].ToString();

                numericUpDown1.Value = decimal.Parse(v.Substring(5, 2));
                numericUpDown2.Value = decimal.Parse(v.Substring(8, 2));
                numericUpDown3.Value = decimal.Parse(v.Substring(0, 4));
                numericUpDown4.Value = decimal.Parse(dt.Rows[0]["PERIOD"].ToString());
               
                
            }
            else
            {
                UPDATE_DATE();

            }

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
            if (period.GET_NUM_ID() == "")
            {
            }
            else if (!bc.exists("SELECT * FROM Period"))
            {

                SQlcommandE(period .sql );
            }
       
            else
            {

                SQlcommandE(sql2 + " WHERE FINANCIAL_YEAR='" + bc.getOnlyString("SELECT FINANCIAL_YEAR FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'") +
                    "' AND PERIOD='" + bc.getOnlyString("SELECT PERIOD FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'") + "'");
            }

            bind();
            F1.bind();


           /* BaseInfo.FrmCurrencyT frm = new FINANCIAL.BaseInfo.FrmCurrencyT();
            frm.LoadAgain();
            frm.bind();
            frm.CYCODE = "RMB";
            frm.CYNAME = "人民币";
            PERIOD PE = new PERIOD();
            frm.FINANCIAL_YEAR = PE.FINANCIAL_YEAR;
            frm.save();*/
            //bc.getcom(" UPDATE CURRENCY_MST SET FINANCIAL_YEAR ='"+FINACIAL_YEAR +"' WHERE CYID='CY001'");

            if (bc.exists("SELECT * FROM CURRENCY_MST"))
            {


            }
            else
            {
                for (int i = 1; i <= 12; i++)
                {
                    SQlcommandE1(sql3, i);
                }
                SQlcommandE2(sql4);
            }
            INITIAL_CONSULENZA FRM = new INITIAL_CONSULENZA();
            FRM.FINANCIAL_YEAR_INITIAL_DATE = FINANCIAL_YEAR_INITIAL_DATE;
            FRM.ShowDialog();
        }
        #region SQlcommandE
        protected void SQlcommandE(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            string varMakerID = "";
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@PEID", SqlDbType.VarChar, 20).Value = period.GET_NUM_ID();
            sqlcom.Parameters.Add("@FINANCIAL_YEAR_INITIAL_DATE", SqlDbType.VarChar, 20).Value = FINANCIAL_YEAR_INITIAL_DATE;
            sqlcom.Parameters.Add("@ACCOUNTING_PERIOD_START_DATE", SqlDbType.VarChar, 20).Value = ACCOUNTING_PERIOD_START_DATE;
            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = numericUpDown3.Value;
            sqlcom.Parameters.Add("@PERIOD", SqlDbType.VarChar, 20).Value = numericUpDown4.Value;
            sqlcom.Parameters.Add("@ACCOUNT_IF_START_USING", SqlDbType.VarChar, 20).Value = "N";
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@IF_CURRENT_ACCOUNTING_PERIOD", SqlDbType.VarChar, 20).Value = 'Y';
            sqlcom.Parameters.Add("@ACCOUNTING_PERIOD_EXPIRATION_DATE", SqlDbType.VarChar, 20).Value = ACCOUNTING_PERIOD_EXPIRATION_DATE;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion

        #region SQlcommandE1
        protected void SQlcommandE1(string sql,int i)
        {
          
          
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            //string varMakerID = "";
            string CYKEY = bc.numYMD(20, 12, "000000000001", "select * from CURRENCY_DET", "CYKEY", "CY");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@CYKEY", SqlDbType.VarChar, 20).Value = CYKEY;
            sqlcom.Parameters.Add("@CYID", SqlDbType.VarChar, 20).Value = "CY0001";
            sqlcom.Parameters.Add("@PERIOD", SqlDbType.VarChar, 20).Value = Convert.ToString(i);
            sqlcom.Parameters.Add("@INITIAL_RATE", SqlDbType.VarChar, 20).Value = 1;
            sqlcom.Parameters.Add("@CLOSING_RATE", SqlDbType.VarChar, 20).Value = 1;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion

        #region SQlcommandE
        protected void SQlcommandE2(string sql)
        {
             string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID = bc.getOnlyString("SELECT EMID FROM USERINFO WHERE USID='" + n2 + "'");
            //string varMakerID = "";
            string CYKEY = bc.numYMD(20, 12, "000000000001", "select * from CURRENCY_DET", "CYKEY", "CY");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            PERIOD PE = new PERIOD();
            sqlcom.Parameters.Add("@CYID", SqlDbType.VarChar, 20).Value = "CY0001";
            sqlcom.Parameters.Add("@CYCODE", SqlDbType.VarChar, 20).Value = "RMB";
            sqlcom.Parameters.Add("@CYNAME", SqlDbType.VarChar, 20).Value = "人民币";
            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = PE.FINANCIAL_YEAR;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = "";
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
            this.Close();
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


        private void UPDATE_DATE()
        {
            
            string v = numericUpDown3.Value.ToString() + "/" + numericUpDown1.Value.ToString() + "/" + numericUpDown2.Value.ToString();
            DateTime d2 = Convert.ToDateTime(v);
            string v1 = d2.ToString("yyyy/MM/dd").Replace("-", "/");
            DateTime d3 = Convert.ToDateTime(v1);
            DateTime d4 = d3.AddMonths(+Convert.ToInt32(numericUpDown4.Value - 1));
            label10.Text = d4.ToString("yyyy") + "年" + d4.ToString("MM") + "月" + d4.ToString("dd") + "日";
            FINANCIAL_YEAR_INITIAL_DATE = v1;
            ACCOUNTING_PERIOD_START_DATE = d4.ToString("yyyy") + "/" + d4.ToString("MM") + "/" + d4.ToString("dd");
            DateTime d5 = d4.AddMonths(+1).AddDays(-1);
            ACCOUNTING_PERIOD_EXPIRATION_DATE = d5.ToString("yyyy/MM/dd");
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }



    }
}
