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
    public partial class VOUCHERT : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        private string _ACID;
        public string ACID
        {
            set { _ACID = value; }
            get { return _ACID; }

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
        protected int i, j;
        protected int M_int_judge, t;
        basec bc = new basec();
        CVOUCHER vou = new CVOUCHER();
        ExcelToCSHARP etc = new ExcelToCSHARP();

        BASE_INFO.CURRENCY cur = new CSPSS.BASE_INFO.CURRENCY();
        VOUCHER F1 = new VOUCHER();
        public VOUCHERT()
        {
            InitializeComponent();
        }
        public VOUCHERT(VOUCHER Frm)
        {
            InitializeComponent();
            F1 = Frm;
        }
        private void VOUCHERT_Load(object sender, EventArgs e)
        {
            textBox1.Text = ACID;
            bind();
           DataTable  dtx=bc.getdt("SELECT VOUCHER_DATE FROM VOUCHER_MST WHERE VOID='"+ACID+"'");
            if (dtx.Rows .Count >0)
            {
                dateTimePicker1.Text = dtx.Rows[0]["VOUCHER_DATE"].ToString();


            }
               
            else
            {
                dateTimePicker1.Text = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");

            }
         
            try
            {
         
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
            this.WindowState = FormWindowState.Maximized;
            Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
            t1.BackColor = c;
            t2.BackColor = c;
            t3.BackColor = c;
            t4.BackColor = c;
        
        }

        #region bind
        private void bind()
        {
            DataTable dtx = basec.getdts(vou.getsql + " where A.VOID='" + textBox1.Text + "' ORDER BY  A.VOKEY ASC ");
                if (dtx.Rows.Count > 0)
                {
                   
                   
                    dt = vou.GET_TABLEINFO(dtx,1);
                    if (dt.Rows.Count > 0 && dt.Rows.Count < 6)
                    {
                        int n = 6 - dt.Rows.Count;
                        for (int i = 0; i <n; i++)
                        {
                           
                            DataRow dr = dt.NewRow();
                            int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                            dr["项次"] = Convert.ToString(b1 + 1);
                            dr["币别"] = dt.Rows[dt.Rows.Count - 1]["币别"].ToString();
                            dr["汇率"] = decimal.Parse(dt.Rows[dt.Rows.Count - 1]["汇率"].ToString());
                            dt.Rows.Add(dr);
                        }
                    }
                }
                else
                {
                    dt = total1();
                    
                }
           
         
            dataGridView1.DataSource = dt;
            dgvStateControl();
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
              dataGridView1.Columns["项次"].Width =40;
              dataGridView1.Columns["摘要"].Width =200;
              dataGridView1.Columns["会计科目"].Width =200;
              dataGridView1.Columns["币别"].Width =40;
              dataGridView1.Columns["汇率"].Width =60;
              dataGridView1.Columns["单价"].Width =60;
              dataGridView1.Columns["数量"].Width =60;
              dataGridView1.Columns["借方原币金额"].Width =100;
              dataGridView1.Columns["借方本币金额"].Width =100;
              dataGridView1.Columns["贷方原币金额"].Width =100;
              dataGridView1.Columns["贷方本币金额"].Width =100;

            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
        
    
            dataGridView1.Columns["摘要"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["会计科目"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["借方原币金额"].DefaultCellStyle.BackColor = Color.Yellow;
            dataGridView1.Columns["贷方原币金额"].DefaultCellStyle.BackColor = Color.Yellow;

            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["摘要"].ReadOnly = false;
            dataGridView1.Columns["会计科目"].ReadOnly = false;
            dataGridView1.Columns["币别"].ReadOnly = false;
            dataGridView1.Columns["汇率"].ReadOnly = false;
            dataGridView1.Columns["单价"].ReadOnly = false;
            dataGridView1.Columns["数量"].ReadOnly = false;
            dataGridView1.Columns["借方原币金额"].ReadOnly = false;
            dataGridView1.Columns["借方本币金额"].ReadOnly = false;
            dataGridView1.Columns["贷方原币金额"].ReadOnly = false;
            dataGridView1.Columns["贷方本币金额"].ReadOnly = false;
          

        }
        #endregion
     
        #region total1
        private DataTable total1()
        {
            DataTable dtt2 = vou.GetTableInfo();
            for (i = 1; i <= 6; i++)
            {
                DataRow dr = dtt2.NewRow();
                dr["项次"] = i;
                dr["币别"] ="RMB";
                dr["汇率"] = "1";
                //dr["借方原币金额"] = "0";
                dtt2.Rows.Add(dr);
            }
            return dtt2;
        }
        #endregion
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter &&(( !(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn) ))
            {
               
                if (dataGridView1.CurrentCell.ColumnIndex == 7 && 
                    dataGridView1["借方原币金额",dataGridView1.CurrentCell.RowIndex].Value .ToString ()!=null )
                {
                    
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else if (dataGridView1.CurrentCell.ColumnIndex == 9 )
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else
                {

                    SendKeys.SendWait("{Tab}");
                }
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
             
                return true;
            }
            if (keyData == (Keys.F7))
            {

                double_info();
              
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
      
        #region juage()
        private bool juage()
        {
            bool b = false;
            for (int k = 0; k <dt.Rows .Count ; k++)
            {
                if (juage(k))
                {
                    b = true;
                    break;
                }
            }
            return b;
        }
        #endregion

        #region juage1()
        private int juage_ABSTRACT_NOEMPTY()
        {
            //int a = dataGridView1.CurrentCell.ColumnIndex;
            //int a1 = dataGridView1.CurrentCell.RowIndex;
           int n=0 ;
            for (int k =dt.Rows .Count -1; k >=0; k--)
            {

                if (dt.Rows[k]["借方原币金额"].ToString() != "" && dt.Rows[k]["贷方原币金额"].ToString() == "" 
                    || dt.Rows[k]["借方原币金额"].ToString() == "" && dt.Rows[k]["贷方原币金额"].ToString() != "")
                {
                    n=k;
                    break;

                }
            }
            return n;

        }
        #endregion
        #region juage()
        private bool juage(int k)
        {
            bool b = false;
           
                string v1 = dt.Rows[k]["摘要"].ToString();
                string v2 =bc.REMOVE_NAME(dt.Rows[k]["会计科目"].ToString());
                string v3 = dt.Rows[k]["币别"].ToString();
                string v4 = dt.Rows[k]["汇率"].ToString();
                string v5 = dt.Rows[k]["单价"].ToString();
                string v6 = dt.Rows[k]["数量"].ToString();
                string v7 = dt.Rows[k]["借方原币金额"].ToString();
                string v8 = dt.Rows[k]["贷方原币金额"].ToString();
                if (v2=="" && v7=="" && v8=="")
                {
                
                }
                else  if (bc.CheckKeyInValueIfNoExistsOrEmpty("ACCOUNTANT_COURSE", "ACCODE", v2, "会计科目"))
                {
                  
                    b = true;
                }
                else if (v2 != "" && v7 == "" && v8 == "")
                {
                    b = true;
                    MessageBox.Show("科目代码不为空时需输入相关金额！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (etc.CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", v2, "会计科目","存在明细科目，需使用明细科目记帐！")==1)
                {
                    b = true;
                }
                else if (bc.CheckKeyInValueIfNoExistsOrEmpty("CURRENCY_MST", "CYCODE", v3, "币别"))
                {
                    b = true;
                }
                else if (bc.CheckKeyInValueIfNoDigitOrEmpty(v4, "汇率"))
                {
                    b = true;

                }
                else if (v7 != "" && v8 != "")
                {
                    b = true;
                    MessageBox.Show("借方原币金额与贷方原币金额同行只能输入一方！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else if (bc.IFEXISTS_LOWERCASE(v3))
                {
                    dt.Rows[k]["币别"] = bc.LOWERCASE_TO_CAPITAL(v3);
                }
               
            return b;
        }
        #endregion
        #region dgvDataSourceChanged
        private void dataGridView1_DataSourceChanged(object sender, EventArgs e)
        {
            int i;
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].ValueType.ToString() == "System.Decimal")
                {
                    
                    dataGridView1.Columns[i].DefaultCellStyle.Format = "#0.00";
                    dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                }
              
            }
            if (dataGridView1.Columns["汇率"].ValueType.ToString() == "System.Decimal")
            {
                dataGridView1.Columns["汇率"].DefaultCellStyle.Format = "#0.0000";
                dataGridView1.Columns["汇率"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
            }
        }
        #endregion
        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("只能输入数字！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        #region btnExcelPrint
        private void btnExcelPrint_Click(object sender, EventArgs e)
        {
           /* try
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

            }*/
        }
        #endregion
        private void ClearText()
        {
          
            dateTimePicker1.Text = "";
            t1.Text = "";
            t2.Text = "";
            t3.Text = "";
            t4.Text = "";
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Focus();
            PERIOD period = new PERIOD(textBox1.Text);
            PERIOD periodo = new PERIOD();
            ACCOUNTING_PERIOD_START_DATE = bc.getOnlyString("SELECT ACCOUNTING_PERIOD_START_DATE FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'");
            ACCOUNTING_PERIOD_EXPIRATION_DATE = bc.getOnlyString("SELECT ACCOUNTING_PERIOD_EXPIRATION_DATE FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'");
            string v1 = ACCOUNTING_PERIOD_EXPIRATION_DATE + " 23:59:59";
            DateTime date1 = Convert.ToDateTime(ACCOUNTING_PERIOD_START_DATE);
            DateTime date2 = Convert.ToDateTime(v1);
            DateTime date3 = Convert.ToDateTime(dateTimePicker1.Value);
            
            dgvfoucs();
           
            string v3 = bc.getOnlyString("SELECT FINANCIAL_YEAR FROM PERIOD");
          
            if (string.IsNullOrEmpty(v3))
            {
                MessageBox.Show("做凭证前需维护会计年度信息！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
         
            else if (period.JUAGE_IF_NO_CURRENT_ACCOUNTING_PERIOD)
            {
                MessageBox.Show(period.ERROW);
            }
            else if (bc.exists("SELECT * FROM VOUCHER_MST WHERE VOID='" + textBox1.Text + "' AND STATUS='CARRY'"))
            {
                MessageBox.Show("结转凭证不能修改！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else if (textBox1.Text == "")
            {

                MessageBox.Show("凭证号不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (date3 < date1 || date3 > date2)
            {
                MessageBox.Show("凭证日期:" + date3.ToString ("yyy/MM/dd")+ "不在当前期间" + ACCOUNTING_PERIOD_START_DATE + "~" + ACCOUNTING_PERIOD_EXPIRATION_DATE + "中！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else if (juage2())
            {

                
            }

            else
            {
                DataTable dtx = bc.GET_NOEMPTY_ROW_COURSE_DT(dt);
                vou.VOUCHER_DATE = dateTimePicker1.Value.ToString("yyyy/MM/dd").Replace("-", "/");
                string da = bc.getOnlyString("SELECT ACCOUNTING_PERIOD_START_DATE FROM PERIOD WHERE IF_CURRENT_ACCOUNTING_PERIOD='Y'");
                DateTime de = Convert.ToDateTime(da);
                DateTime de1 = de.AddMonths(+1).AddDays(-1);
                vou.ACCOUNTING_PERIOD_EXPIRATION_DATE = de1.ToString("yyyy/MM/dd");
                vou.save("VOUCHER_MST", "VOUCHER_DET", "VOID", textBox1.Text, dtx, "OPEN");
                LoadAgain();
                F1.Bind();
            }
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region juage2()
        private bool juage2()
        {
            bool b = false;
            string v5 = dt.Compute("sum(借方原币金额)","").ToString();
            string v6 = dt.Compute("sum(贷方原币金额)","").ToString();
            string v7 = dt.Compute("sum(借方本币金额)","").ToString();
            string v8 = dt.Compute("sum(贷方本币金额)","").ToString();
            decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0;
            if (!string.IsNullOrEmpty(v5))
            {
                d1 = decimal.Parse(v5);
            }
            if (!string.IsNullOrEmpty(v6))
            {
                d2 = decimal.Parse(v6);
            }
            if (!string.IsNullOrEmpty(v7))
            {
                d3 = decimal.Parse(v7);
            }
            if (!string.IsNullOrEmpty(v8))
            {
                d4= decimal.Parse(v8);
            }
            if (juage())
            {
                b = true;
              
            }
            else if (d3 != d4)
            {
                b = true;
                MessageBox.Show("借方本币金额不等于贷方本币金额！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else if (juage_ABSTRACT_NOEMPTY() != 0)
            {
                if (dt.Rows[juage_ABSTRACT_NOEMPTY ()]["摘要"].ToString() == "")
                {
                    b = true;
                    MessageBox.Show("项次" + dt.Rows[juage_ABSTRACT_NOEMPTY()]["项次"].ToString() + "摘要不能为空！",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                  
                }

            }
            return b;

        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
      
        private void btnDel_Click(object sender, EventArgs e)
        {
            
            try
            {
                PERIOD period = new PERIOD(textBox1.Text);
                if (period.JUAGE_IF_NO_CURRENT_ACCOUNTING_PERIOD)
                {
                    MessageBox.Show(period.ERROW);
                }
                else if (MessageBox.Show("确定要删除该条凭证吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    basec.getcoms("DELETE VOUCHER_MST WHERE VOID='" + textBox1.Text + "'");
                    basec.getcoms("DELETE VOUCHER_DET WHERE VOID='" + textBox1.Text + "'");
                    bind();
                    ClearText();
                    textBox1.Text = "";
                    F1.Bind();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #region dgvCellEndEdit
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int a = dataGridView1.CurrentCell.ColumnIndex;
            int b = dataGridView1.CurrentCell.RowIndex;
            int c = dataGridView1.Columns.Count - 1;
            int d = dataGridView1.Rows.Count - 1;
            if (a == 2)
            {
                if (!string.IsNullOrEmpty(dt.Rows[b]["会计科目"].ToString()))
                {
                    dt2 = bc.getdt(etc.getsql + " WHERE A.ACCODE='" + dt.Rows[b]["会计科目"].ToString() + "'");
                    if (dt2.Rows.Count > 0)
                    {

                        dt.Rows[b]["会计科目"] = dt.Rows[b]["会计科目"].ToString() + " " + etc.GetLastCourseAnd_CurrentCourseName(dt.Rows[b]["会计科目"].ToString());
                    }
                }

            }
            if (a == 3)/*CURRENCY*/
            {

                PERIOD period = new PERIOD();
                dt2 = bc.getdt(cur.GETSQL + " WHERE B.CYCODE='" + dt.Rows[b]["币别"].ToString() +
                    "' AND B.FINANCIAL_YEAR='" + period.FINANCIAL_YEAR + "' AND A.PERIOD='" + period.GETPERIOD + "'");
                if (dt2.Rows.Count > 0)
                {
                    if (bc.IFEXISTS_LOWERCASE(dt.Rows[b]["币别"].ToString()))
                    {
                        dt.Rows[b]["币别"] = bc.LOWERCASE_TO_CAPITAL(dt.Rows[b]["币别"].ToString());
                    }
                    dt.Rows[b]["汇率"] = dt2.Rows[0]["期初汇率"].ToString();
                }

            }
            try
            {

       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        #endregion
        #region dgvDoubleClick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int currentrowsindex = dataGridView1.CurrentCell.RowIndex;
            int currentcolumnindex = dataGridView1.CurrentCell.ColumnIndex;
            if (currentcolumnindex == 2)
            {
                
                BASE_INFO.ACCOUNTANT_COURSE frm = new BASE_INFO.ACCOUNTANT_COURSE();
                frm.a5();
                frm.ShowDialog();
                dataGridView1["会计科目", currentrowsindex].Value = frm.ACCODE;
                dataGridView1.CurrentCell = dataGridView1["币别", dataGridView1.CurrentCell.RowIndex];
                 
            }
            if (currentcolumnindex == 3)
            {
                BASE_INFO.CURRENCY frm = new CSPSS.BASE_INFO.CURRENCY();
                frm.a5();
                frm.ShowDialog();
                dataGridView1["币别", currentrowsindex].Value = frm.CYCODE;
                dataGridView1["汇率", currentrowsindex].Value = frm.EXCHANGE_RATE;
                dataGridView1.CurrentCell = dataGridView1["单价", dataGridView1.CurrentCell.RowIndex];
            }
            try
            {
    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        }
        #endregion
        private void double_info()
        {

            BASE_INFO.ACCOUNTANT_COURSE frm = new BASE_INFO.ACCOUNTANT_COURSE();
            frm.a5();
            frm.ShowDialog();
            DataGridViewRow dgvr = dataGridView1.CurrentRow;
            int j = dataGridView1.CurrentCell.ColumnIndex;
            if (dataGridView1.Columns[j].Name == "会计科目")
            {
                dgvr.Cells["会计科目"].Value = frm.ACCODE;
                dataGridView1.CurrentCell = dataGridView1["币别", dataGridView1.CurrentCell.RowIndex];
            } 
        }

        #region dgvCellEnter
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            int a = dataGridView1.CurrentCell.ColumnIndex;
            int b = dataGridView1.CurrentCell.RowIndex;
            int c = dataGridView1.Columns.Count - 1;
            int d = dataGridView1.Rows.Count - 1;


            if (a == c && b == d)
            {
                if (dt.Rows.Count >= 6)
                {

                    DataRow dr = dt.NewRow();
                    int b1 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["项次"].ToString());
                    dr["项次"] = Convert.ToString(b1 + 1);
                    dr["币别"] = dt.Rows[dt.Rows.Count - 1]["币别"].ToString();
                    dr["汇率"] = decimal.Parse(dt.Rows[dt.Rows.Count - 1]["汇率"].ToString());
                    dt.Rows.Add(dr);
                }

            }
            dgvfoucs();
            try
            {
         
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
        #region ask
        private void ask(int k)
        {
            int n = k;
            decimal v1 = decimal.Parse(dt.Rows[k]["汇率"].ToString());
            decimal v2=0, v3=0;
            if (!string.IsNullOrEmpty(dt.Rows[k]["借方原币金额"].ToString()))
            {
                v2 = decimal.Parse(dt.Rows[k]["借方原币金额"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[k]["贷方原币金额"].ToString()))
            {
                v3 = decimal.Parse(dt.Rows[k]["贷方原币金额"].ToString());
            }
            if (v2 > 0)
            {

                dt.Rows[n]["借方本币金额"] = v1 * v2;
            }
            if (v3 > 0)
            {

                dt.Rows[n]["贷方本币金额"] = v1 * v3;
            }
      
            ask1();
        }
        #endregion
        #region ask1
        private void ask1()
        {
            t1.Text = "";
            t2.Text = "";
            t3.Text = "";
            t4.Text ="";
            string v5 = dt.Compute("sum(借方原币金额)", "").ToString();
            string v6 = dt.Compute("sum(贷方原币金额)", "").ToString();
            string v7 = dt.Compute("sum(借方本币金额)", "").ToString();
            string v8 = dt.Compute("sum(贷方本币金额)", "").ToString();
            if (!string.IsNullOrEmpty(v5))
            {
                t1.Text = string.Format("{0:F2}", Convert.ToDouble(v5));
            
            }
            if (!string.IsNullOrEmpty(v7))
            {
                
                t3.Text = string.Format("{0:F2}", Convert.ToDouble(v7));
            }
            if (!string.IsNullOrEmpty(v6))
            {
                t2.Text = string.Format("{0:F2}", Convert.ToDouble(v6));
             
            }
            if (!string.IsNullOrEmpty(v8))
            {
                t4.Text = string.Format("{0:F2}", Convert.ToDouble(v8));
            }
        }
        #endregion
        #region dgvCellValidating
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.ColumnIndex == 2 && bc.CheckKeyInValueIfNoExists("ACCOUNTANT_COURSE", "ACCODE",
                 bc.REMOVE_NAME(e.FormattedValue.ToString()), "会计科目"))
            {

                e.Cancel = true;
            }
            else if (e.ColumnIndex == 2 && e.FormattedValue.ToString() != "" &&
                 etc.CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", bc.REMOVE_NAME(e.FormattedValue.ToString()),
                 "会计科目", "存在明细科目，需使用明细科目记帐！") == 1)
            {

                e.Cancel = true;
            }
            else if (e.ColumnIndex == 3 && bc.CheckKeyInValueIfNoExistsOrEmpty("CURRENCY_MST", "CYCODE", e.FormattedValue.ToString(), "币别"))
            {

                e.Cancel = true;
            }
            else if (e.ColumnIndex == 4 && bc.CheckKeyInValueIfNoDigitOrEmpty(e.FormattedValue.ToString(), "汇率"))
            {

                e.Cancel = true;
            }
            else if (e.ColumnIndex == 5 && bc.yesno(e.FormattedValue.ToString()) == 0)
            {
                e.Cancel = true;
                MessageBox.Show("单价只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }
            else if (e.ColumnIndex == 6 && bc.yesno(e.FormattedValue.ToString()) == 0)
            {
                e.Cancel = true;
                MessageBox.Show("数量只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }
            try
            {
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
        private void dgvfoucs()
        {
            
            for (i = 0; i < dt.Rows .Count ; i++)
            {
                ask(i);
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          
        }

        #region loadagain
        private void LoadAgain()
        {
            ClearText();
            dt = total1();
            dataGridView1.DataSource = dt;
            dgvStateControl();

            string a1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_MST", "VOID", "VO");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                textBox1.Text = a1;
            }
        }
        #endregion
        private void TSMI_Click(object sender, EventArgs e)
        {
            dgvclear(dataGridView1.CurrentCell.RowIndex);
            
        }
        private void dgvclear(int r)
        {
            
            dt.Rows[r]["摘要"] = "";
            dt.Rows[r]["会计科目"] = null;
            //dt.Rows[r]["币别"] = "";

            //dt.Rows[r]["汇率"] = DBNull.Value;
            dt.Rows[r]["单价"] = "";
            dt.Rows[r]["数量"] = "";
            dt.Rows[r]["借方原币金额"] = DBNull.Value;
            dt.Rows[r]["借方本币金额"] = DBNull.Value;
            dt.Rows[r]["贷方原币金额"] = DBNull.Value;
            dt.Rows[r]["贷方本币金额"] = DBNull.Value;
            btnSave.Focus();
        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
            PERIOD period = new PERIOD(textBox1.Text);
            if (period.JUAGE_IF_NO_CURRENT_ACCOUNTING_PERIOD)
            {
                MessageBox.Show(period.ERROW);
            }
            else
            {
                dgvclear(dataGridView1.CurrentCell.RowIndex);
            }
        }

        private void btnAllSelect_Click(object sender, EventArgs e)
        {
            PERIOD period = new PERIOD(textBox1.Text);
            if (period.JUAGE_IF_NO_CURRENT_ACCOUNTING_PERIOD)
            {
                MessageBox.Show(period.ERROW);
            }
            else 
            {
                for (i = 0; i < dt.Rows.Count; i++)
                {
                    dgvclear(i);
                }
            }
        }

        private void dataGridView1_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
         
            int r=dataGridView1.CurrentCell.RowIndex;
            if (dataGridView1["借方原币金额", r].Value.ToString() != "" && dataGridView1["贷方原币金额", r].Value.ToString() != "")
            {
                e.Cancel = true;
                MessageBox.Show("借方原币金额与贷方原币金额同行只能输入一方！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void 提取科目F7ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double_info();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

   
   


    }
}
