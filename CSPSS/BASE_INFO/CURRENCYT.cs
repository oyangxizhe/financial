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
    public partial class CURRENCYT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtd = new DataTable();
        DataTable dt3 = new DataTable();
        DataTable dt4 = new DataTable();
        DataTable dt5 = new DataTable();
        DataTable dt6 = new DataTable();
        DataTable dtx2 = new DataTable();
        DataTable dtx3 = new DataTable();
        DataTable dtnn = new DataTable();
        basec boperate = new basec();
        //  C23.BaseClass.OperateAndValidate opAndvalidate = new C23.BaseClass.OperateAndValidate();
        protected int i, j;
        public static string[] inputTextDataWare = new string[] { null, null, null, null, null, null, null, null, null };
        public static string[] inputTextDataStorage = new string[] { "" };
        public static string[] inputTextDataLocation = new string[] { "" };
        public static string[] inputgetSEName = new string[] { "" };
        public static string[] str1 = new string[] { "" };
        public static string[] str2 = new string[] { "", "" };
        public static string[] str4 = new string[] { "" };
        public static string[] str6 = new string[] { "", "", "", "", "" };
        public static string[] str7 = new string[] { "" };
        public static string[] str8 = new string[] { "", "", "", "" };
        public static string[] data1 = new string[] { "" };
        public static string[] data2 = new string[] { "", "", "", "" };
        string[] a = new string[] { "", "加急" };
        protected string sql = @"
SELECT
A.CYKEY AS 索引,
A.CYID AS 币别编号,
B.CYCODE AS 币别,
B.CYNAME AS 名称,
B.FINANCIAL_YEAR AS 会计年度,
A.PERIOD AS 期间,
A.INITIAL_RATE AS 期初汇率,
A.CLOSING_RATE AS 期末汇率,
B.MAKERID AS 制单人工号,
B.DATE AS  制单日期 FROM CURRENCY_DET A 
LEFT JOIN CURRENCY_MST B ON  A.CYID=B.CYID";

        protected int M_int_judge, t;
        string CYKEY;
 
        private string _CYID;
        public string CYID
        {
            set { _CYID = value; }
            get { return _CYID; }

        }
        private string _CYCODE;
        public string CYCODE
        {
            set { _CYCODE = value; }
            get { return _CYCODE; }

        }
        private string _CYNAME;
        public string CYNAME
        {
            set { _CYNAME = value; }
            get { return _CYNAME; }

        }
        private string _FINANCIAL_YEAR;
        public string FINANCIAL_YEAR
        {
            set { _FINANCIAL_YEAR = value; }
            get { return _FINANCIAL_YEAR; }

        }
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
        basec bc = new basec();
        CCURRENCY currency = new CCURRENCY();

        CURRENCY FRM = new CURRENCY();
        public CURRENCYT()
        {
            InitializeComponent();


        }
        public CURRENCYT(CURRENCY frm)
        {
            InitializeComponent();
            FRM = frm;

        }
        private void CURRENCYT_Load(object sender, EventArgs e)
        {
            if (str1[0] != "")
            {
                textBox1.Text = str1[0];
                str1[0] = "";
            }
            else
            {
                textBox1.Text = str6[0];
                str6[0] = "";
            }
            bind();
            try
            {
          

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #region loadagain
        public void LoadAgain()
        {
            CCURRENCY currencyo = new CCURRENCY();
            if (currencyo .NUMID != "")
            {
                textBox1.Text = currencyo.NUMID;
                CYID = currencyo.NUMID;
            }
        }
        #endregion
        #region bind
        public void bind()
        {
            dt = basec.getdts(sql + " where A.CYID='" + textBox1.Text + "'");
            if (dt.Rows.Count > 0)
            {
                textBox2.Text = dt.Rows[0]["币别"].ToString();
                textBox3.Text = dt.Rows[0]["名称"].ToString();
                //MessageBox.Show(dt.Rows[0]["会计年度"].ToString()+"/01/01");
                DateTime date1 = Convert.ToDateTime(dt.Rows[0]["会计年度"].ToString() + "/01/01");
                dateTimePicker1.Text = Convert.ToString(date1);
                dataGridView1.DataSource = as1();
            }
            else
            {

                dataGridView1.DataSource = total1();
            }
            dgvStateControl();

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

                if (i == 1 || i == 2)
                {
                    dataGridView1.Columns[i].Width = 90;
                }
                else
                {
                    dataGridView1.Columns[i].Width = 60;

                }

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
                if (i == 1 || i == 10 || i == 12 || i == 13 || i == 14)
                {
                    //dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.Yellow;
                }
                if (i == 15)
                {
                    dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.GreenYellow;

                }
            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 1 || i == 2)
                {
                    dataGridView1.Columns[i].ReadOnly = false;
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
            textBox2.BackColor = Color.Yellow;
            textBox3.BackColor = Color.Yellow;

        }
        #endregion
        #region total
        private DataTable total()
        {
            dt = new DataTable();
            //dt.Columns.Add("索引", typeof(string));
            dt.Columns.Add("期间", typeof(string));
            dt.Columns.Add("期初汇率", typeof(decimal));
            dt.Columns.Add("期末汇率", typeof(decimal));
            return dt;
        }
        #endregion
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = total();
            for (i = 1; i <= 12; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["期间"] = Convert.ToString(i);
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region as1()
        private DataTable as1()
        {

            DataTable dtx1 = basec.getdts(sql + " WHERE A.CYID='" + textBox1.Text + "'");
            DataTable dt2 = total();
            foreach (DataRow dr in dtx1.Rows)
            {

                DataRow dr1 = dt2.NewRow();
                //dr1["索引"] = dr["索引"].ToString();
                dr1["期间"] = dr["期间"].ToString();
                dr1["期初汇率"] = dr["期初汇率"].ToString();
                dr1["期末汇率"] = dr["期末汇率"].ToString();
                dt2.Rows.Add(dr1);
            }
            return dt2;
        }
        #endregion

        #region save
        public void save()
        {
           
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v3 = bc.getOnlyString("SELECT CYCODE FROM CURRENCY_MST WHERE  CYID='" + CYID + "'");
            string v4 = bc.getOnlyString("SELECT FINANCIAL_YEAR FROM CURRENCY_MST WHERE  CYID='" + CYID + "'");
            //string varMakerID;
        
            string v1, v2;
            if (!bc.exists("SELECT * FROM CURRENCY_DET WHERE CYID='" + CYID + "'"))
            {

                if (bc.exists("SELECT * FROM CURRENCY_MST WHERE CYCODE='" + CYCODE + "' AND FINANCIAL_YEAR='" + FINANCIAL_YEAR + "'"))
                {

                    MessageBox.Show("币别与会计年度已经存在于系统", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {


                    foreach (DataRow dr in dt.Rows)
                    {
                        CCURRENCY currencyx = new CCURRENCY();
                        CYKEY = currencyx.KEY;
                   
                        if (dr["期初汇率"].ToString() != "")
                        {
                            v1 = dr["期初汇率"].ToString();
                        }
                        else
                        {
                            v1 = Convert.ToString(1);
                        }
                        if (dr["期末汇率"].ToString() != "")
                        {
                            v2 = dr["期末汇率"].ToString();
                        }
                        else
                        {
                            v2 = Convert.ToString(1);
                        }

                        string sql = @"INSERT INTO CURRENCY_DET(CYKEY,CYID,PERIOD,INITIAL_RATE,CLOSING_RATE,YEAR,MONTH,DAY)
VALUES ('" + CYKEY + "','" + CYID + "','" + dr["期间"].ToString() + "','" + v1 + "','" + v2 + "','" + year +
           "','" + month + "','" + day + "')";
                        basec.getcoms(sql);

                    }
                }
            }
            else
            {
                if (v3 != CYCODE || v4 != FINANCIAL_YEAR)
                {
                    if (bc.exists("SELECT * FROM CURRENCY_MST WHERE CYCODE='" + CYCODE + "' AND FINANCIAL_YEAR='" + FINANCIAL_YEAR + "'"))
                    {

                        MessageBox.Show("币别与会计年度已经存在于系统！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        foreach (DataRow dr in dt.Rows)
                        {

                            basec.getcoms("UPDATE CURRENCY_DET SET INITIAL_RATE='" + dr["期初汇率"].ToString() +
                                "',CLOSING_RATE='" + dr["期末汇率"].ToString() +
                                "' FROM CURRENCY_DET A LEFT JOIN CURRENCY_MST B ON A.CYID=B.CYID WHERE "
                                + " B.CYID='" + CYID + "' AND A.PERIOD='" + dr["期间"].ToString() + "'");
                        }
                    }
                }
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {

                        basec.getcoms("UPDATE CURRENCY_DET SET INITIAL_RATE='" + dr["期初汇率"].ToString() +
                            "',CLOSING_RATE='" + dr["期末汇率"].ToString() +
                            "' FROM CURRENCY_DET A LEFT JOIN CURRENCY_MST B ON A.CYID=B.CYID WHERE "
                            + " B.CYID='" + CYID + "' AND A.PERIOD='" + dr["期间"].ToString() + "'");
                    }
                }
            }
            if (!bc.exists("SELECT CYID FROM CURRENCY_DET WHERE CYID='" + CYID + "'"))
            {

                return;
            }
            else if (!bc.exists("SELECT CYID FROM CURRENCY_MST WHERE CYID='" + CYID + "'"))
            {

                SQlcommandE(currency .sql);
            }
            else if (v3 != CYCODE || v4 != FINANCIAL_YEAR)
            {
                if (bc.exists("SELECT * FROM CURRENCY_MST WHERE CYCODE='" + CYCODE + "' AND FINANCIAL_YEAR='" + FINANCIAL_YEAR + "'"))
                {

                    //MessageBox.Show("币别与会计年度已经存在于系统！3", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    SQlcommandE(currency .sqlo + " WHERE CYID='" + CYID + "'");
                }
            }
            else
            {
                SQlcommandE(currency .sqlo  + " WHERE CYID='" + CYID + "'");

            }
         
        }
        #endregion


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
            sqlcom.Parameters.Add("@CYID", SqlDbType.VarChar, 20).Value = CYID;
            sqlcom.Parameters.Add("@CYCODE", SqlDbType.VarChar, 20).Value = CYCODE;
            sqlcom.Parameters.Add("@CYNAME", SqlDbType.VarChar, 20).Value = CYNAME;
            sqlcom.Parameters.Add("@FINANCIAL_YEAR", SqlDbType.VarChar, 20).Value = FINANCIAL_YEAR;
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
        #region juage()
        private bool juage()
        {
            bool b = false;
            for (int k = 0; k < dt.Rows.Count; k++)
            {
                string v1 = dt.Rows[k][1].ToString();
                string v2 = dt.Rows[k][2].ToString();

                if (textBox2.Text == "")
                {

                    b = true;
                    MessageBox.Show("币别不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                }
                else if (textBox3.Text == "")
                {
                    b = true;
                    MessageBox.Show("名称不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;


                }
                else if (bc.yesno(v1) == 0)
                {
                    b = true;
                    MessageBox.Show(v1 + " " + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                }
                else if (bc.yesno(v2) == 0)
                {
                    b = true;
                    MessageBox.Show(v2 + " " + "只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                }


            }
            return b;

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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            ClearText();
            LoadAgain();
            dataGridView1.DataSource = total1();
        }
        private void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            dateTimePicker1.Text = "";

        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Focus();
            int currentcolumnindex = dataGridView1.CurrentCell.ColumnIndex;
            int currentrowindex = dataGridView1.CurrentCell.RowIndex;
            int columncount = dataGridView1.Columns.Count - 1;
            int rowcount = dataGridView1.Rows.Count - 1;
            try
            {
                if (juage())
                {


                }
                else if (CYKEY == "Exceed Limited")
                {

                    MessageBox.Show("编码超出限制!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    CYID = textBox1.Text;
                    CYCODE = textBox2.Text;
                    CYNAME = textBox3.Text;
                    FINANCIAL_YEAR = dateTimePicker1.Text;
                    save();
                    bind();
                    FRM.Bind();
                }
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

        private void btnDel_Click(object sender, EventArgs e)
        {

            try
            {

                if (MessageBox.Show("确定要删除该条信息吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    
                    if (bc.exists("SELECT * FROM ACCOUNTANT_COURSE WHERE CYID='" + textBox1.Text + "'"))
                    {
                        MessageBox.Show("该币别已经存在科目信息中，不允许删除!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (bc.exists("SELECT * FROM VOUCHER_DET WHERE CYID='" + textBox1.Text + "'"))
                    {
                        MessageBox.Show("该币别已经存在凭证信息中，不允许删除!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        basec.getcoms("DELETE CURRENCY_MST WHERE CYID='" + textBox1.Text + "'");
                        basec.getcoms("DELETE CURRENCY_DET WHERE CYID='" + textBox1.Text + "'");
                        bind();
                        ClearText();
                        textBox1.Text = "";
                        FRM.Bind();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
    }
}
