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


namespace FINANCIAL.VOUCHER_MANAGE
{
    public partial class FrmDETAIL_SEARCH : Form
    {
        DataTable dt = new DataTable();
        protected int i, j;
        protected string sql = @"SELECT * FROM PERIOD";
        protected int M_int_judge, t;
        basec bc = new basec();
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
        string sql1 = @"INSERT INTO Period(PEID,
FINANCIAL_YEAR,
PERIOD,
MAKERID,
DATE
) VALUES 

(
@PEID,
@FINANCIAL_YEAR,
@PERIOD,
@MAKERID,
@DATE
)

";
        string sql2 = @"UPDATE Period SET 
PEID=@PEID,
FINANCIAL_YEAR=@FINANCIAL_YEAR,
PERIOD=@PERIOD,
MAKERID=@MAKERID,
DATE=@DATE
";
        public FrmDETAIL_SEARCH()
        {
            InitializeComponent();
           
        }
        private void FrmPeriodT_Load(object sender, EventArgs e)
        {
           
            bind();
            try
            {
                
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
            Color c1 = System.Drawing.ColorTranslator.FromHtml("#990033");
            label4.ForeColor = c1;
            DataTable dtx1 = bc.getdt("SELECT CYCODE FROM CURRENCY_MST ");
            foreach (DataRow dr in dtx1.Rows)
            {

                comboBox3.Items.Add(dr["CYCODE"].ToString());
            }
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            
        }

        #region bind
        public  void bind()
        {
            numericUpDown1.Maximum = 5000;
            numericUpDown1.Value = Convert.ToInt32(DateTime.Now.Date.ToString("yyyy"));
            numericUpDown1.Minimum = 1900;
            numericUpDown1.TextAlign = HorizontalAlignment.Center;

            numericUpDown2.Value = Convert.ToInt32(DateTime.Now.Date.ToString("MM"));
            numericUpDown2.Maximum = 12;
            numericUpDown2.Minimum = 1;
            numericUpDown2.TextAlign = HorizontalAlignment.Center;

           

            numericUpDown4.Value = Convert.ToInt32(DateTime.Now.Date.ToString("MM"));
            numericUpDown4.Maximum = 12;
            numericUpDown4.Minimum = 1;
            numericUpDown4.TextAlign = HorizontalAlignment.Center;
            dt = basec.getdts(sql);
            if (dt.Rows.Count > 0)
            {
               
                DateTime date1 = Convert.ToDateTime(dt.Rows[0]["FINANCIAL_YEAR"].ToString() + "/01/01");
                FINACIAL_YEAR = dt.Rows[0]["FINANCIAL_YEAR"].ToString();
                PERIOD = dt.Rows[0]["PERIOD"].ToString();
            }
          
          
        }
        #endregion
        #region save
        protected void save()
        {
           
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            //string varMakerID;
            if (!bc.exists("SELECT * FROM Period"))
            {

                SQlcommandE(sql1);
            }
            else
            {

                SQlcommandE(sql2+" WHERE FINANCIAL_YEAR='"+bc.getOnlyString ("SELECT FINANCIAL_YEAR FROM PERIOD") +
                    "' AND PERIOD='" + bc.getOnlyString("SELECT PERIOD FROM PERIOD") + "'");
            }
            
            bind();
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
                sqlcom.Parameters.Add("@PEID", SqlDbType.VarChar, 20).Value = "PEID14040017";
                sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = varMakerID;
                sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
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
 
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            label4.Text = "";
            if (juage())
            {

            }
            else if (!bc.exists("SELECT * FROM PERIOD WHERE FINANCIAL_YEAR='" + numericUpDown1.Value + "'"))
            {
                label4.Text = "会计年度不存在！";
            }
            else if (!bc.exists("SELECT * FROM PERIOD WHERE PERIOD='" + numericUpDown2.Value + "'"))
            {

                label4.Text = "起始帐期不存在！";
              
            }
            else if (!bc.exists("SELECT * FROM PERIOD WHERE PERIOD='" + numericUpDown4.Value + "'"))
            {

                label4.Text = "截止帐期不存在！";
            }
            else
            {
                this.Close();
                FrmDETAIL_ACCOUNT FRM = new FrmDETAIL_ACCOUNT();
                string v1 = bc.getOnlyString("SELECT ACCOUNTING_PERIOD_START_DATE FROM PERIOD WHERE FINANCIAL_YEAR='" + numericUpDown1.Value +
                    "' AND PERIOD='" + numericUpDown2.Value + "'");
                string v2 = bc.getOnlyString("SELECT ACCOUNTING_PERIOD_EXPIRATION_DATE FROM PERIOD WHERE FINANCIAL_YEAR='" + numericUpDown1.Value +
               "' AND PERIOD='" + numericUpDown4.Value + "'");

                FRM.ACCOUNTING_PERIOD_START_PERIOD = v1;
                FRM.ACCOUNTING_PERIOD_EXPIRATION_PERIOD = v2;
                FRM.ShowDialog();
            }
           
            /*
          ;*/
        }
        private void UPDATE_DATE()
        {
            string v = numericUpDown1.Value.ToString() + "/" + numericUpDown2.Value.ToString();
            DateTime d2 = Convert.ToDateTime(v);
            string v1 = d2.ToString("yyyy/MM/dd").Replace("-", "/");
            DateTime d3 = Convert.ToDateTime(v1);
     
           
           
        }
        #region juage()
        private bool juage()
        {
            bool b = false;
            int i1 = Convert.ToInt32 (numericUpDown1.Value);
            int i2 = Convert.ToInt32 (numericUpDown2.Value);
            int i3 = Convert.ToInt32(numericUpDown4.Value);
            if(i3<i2)
            {
                label4.Text = "截止期间需大于或等于起始期间！";
                b = true;
            }
            return b;
        }
        #endregion
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            UPDATE_DATE();
        }
    }
}
