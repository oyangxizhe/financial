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
    public partial class INITIAL_CONSULENZA : Form
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
        private string _FINANCIAL_YEAR_INITIAL_DATE;
        public string FINANCIAL_YEAR_INITIAL_DATE
        {

            set { _FINANCIAL_YEAR_INITIAL_DATE = value; }
            get { return _FINANCIAL_YEAR_INITIAL_DATE; }

        }

        private string _EXPIRATION_DATE;
        public string EXPIRATION_DATE
        {

            set { _EXPIRATION_DATE = value; }
            get { return _EXPIRATION_DATE; }

        }
 
        protected int i, j;
        protected int M_int_judge, t;
        basec bc = new basec();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        CVOUCHER vou = new CVOUCHER();
        BASE_INFO.CURRENCY cur = new CSPSS.BASE_INFO.CURRENCY();
  
        public INITIAL_CONSULENZA()
        {
            InitializeComponent();
        }

        private void INITIAL_CONSULENZA_Load(object sender, EventArgs e)
        {

          
            bind();
            Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
            Color c1 = System.Drawing.ColorTranslator.FromHtml("#990033");
            Color c2 = System.Drawing.ColorTranslator.FromHtml("#008000");

            t3.BackColor = c;
            t4.BackColor = c;
            t5.BackColor = c;
            t6.BackColor = c;
            label1.ForeColor = c1;
            label2.ForeColor = c2;
          
            //this.WindowState = FormWindowState.Maximized;
            dataGridView1.CurrentCell = dataGridView1[5, 0];

            try
            {

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        
        }
   

        #region bind
        private void bind()
        {
            DataTable dtx = vou.GET_TABLEINFO_INITIAL();
            if (dtx.Rows.Count > 0)
            {
                dt = dtx;
               
            }
            else
            {
                dt = vou.GET_TABLEINFO_INITIAL_O();

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
            dataGridView1.Columns["项次"].Width = 40;
            dataGridView1.Columns["科目代码"].Width = 100;
            dataGridView1.Columns["科目名称"].Width = 200;
            dataGridView1.Columns["年初借方"].Width = 100;
            dataGridView1.Columns["年初贷方"].Width = 100;
            dataGridView1.Columns["累计借方"].Width = 100;
            dataGridView1.Columns["累计贷方"].Width = 100;
            dataGridView1.Columns["方向"].Width = 40;
            dataGridView1.Columns["期初借方"].Width = 100;
            dataGridView1.Columns["期初贷方"].Width = 100;

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
            Color c = System.Drawing.ColorTranslator.FromHtml("#f4f1fc");

            dataGridView1.Columns["年初借方"].DefaultCellStyle.BackColor = c;
            dataGridView1.Columns["年初贷方"].DefaultCellStyle.BackColor = c;

            dataGridView1.Columns["项次"].ReadOnly = true;
            dataGridView1.Columns["科目代码"].ReadOnly = true;
            dataGridView1.Columns["科目名称"].ReadOnly = true;
            dataGridView1.Columns["年初借方"].ReadOnly = true;
            dataGridView1.Columns["年初贷方"].ReadOnly = true;
            dataGridView1.Columns["累计借方"].ReadOnly = false;
            dataGridView1.Columns["累计贷方"].ReadOnly = false;
            dataGridView1.Columns["方向"].ReadOnly = true;
            dataGridView1.Columns["期初借方"].ReadOnly = false;
            dataGridView1.Columns["期初贷方"].ReadOnly = false;

        }
        #endregion
     
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                 !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {
                if (dataGridView1.CurrentCell.ColumnIndex == 6)
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");

                }
                else if (dataGridView1.CurrentCell.ColumnIndex == 9)
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
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
            return base.ProcessCmdKey(ref msg, keyData);
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
 
        private void btnSave_Click(object sender, EventArgs e)
        {
        
           
            label1.Text = "";
            btnSave.Focus();
            for (i = 0; i < dt.Rows.Count; i++)
            {
                ask(i);
            }
            string v3 = bc.getOnlyString("SELECT FINANCIAL_YEAR FROM PERIOD");
            if (juage2())
            {
             

            }
            else
            {
                DataTable dty = vou.GET_TABLEINFO_INITIAL();
                if (dty.Rows.Count > 0)
                {
                    bc.getcom("DELETE VOUCHER_DET");/* aready do initialize*/
                    bc.getcom("DELETE VOUCHER_MST");

                }
                string a1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_MST", "VOID", "VO");
                if (a1 == "Exceed Limited")
                {
                    MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    DataTable dtx = bc.GET_NOEXISTS_EMPTY_ROW_DT(dt, "",
                        "累计借方 IS NOT NULL OR 累计贷方 IS NOT NULL OR 期初借方 IS NOT NULL OR 期初贷方 IS NOT NULL");

                    string da = bc.getOnlyString("SELECT FINANCIAL_YEAR_INITIAL_DATE FROM PERIOD ");
                    DateTime de = Convert.ToDateTime(da);
                    DateTime de1 = de.AddMonths(+1).AddDays(-1);
                    vou.FINANCIAL_YEAR_INITIAL_DATE = FINANCIAL_YEAR_INITIAL_DATE;
                    vou.VOUCHER_DATE = FINANCIAL_YEAR_INITIAL_DATE;
                    vou.ACCOUNTING_PERIOD_EXPIRATION_DATE = de1.ToString("yyyy/MM/dd");
                    vou.save("VOUCHER_MST", "VOUCHER_DET", "VOID", a1, dtx, "INITIAL");


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
        #region juage2()
        private bool juage2()
        {
            bool b = false;
            string v1 = dt.Compute("sum(年初借方)","").ToString();
            string v2 = dt.Compute("sum(年初贷方)","").ToString();
            string v3 = dt.Compute("sum(累计借方)","").ToString();
            string v4 = dt.Compute("sum(累计贷方)","").ToString();
            string v5 = dt.Compute("sum(期初借方)", "").ToString();
            string v6 = dt.Compute("sum(期初贷方)", "").ToString();
            decimal d1 = 0, d2 = 0, d3 = 0, d4 = 0,d5=0,d6=0;
            if (!string.IsNullOrEmpty(v1))
            {
                d1 = decimal.Parse(v1);
            }
            if (!string.IsNullOrEmpty(v2))
            {
                d2 = decimal.Parse(v2);
            }
            if (!string.IsNullOrEmpty(v3))
            {
                d3 = decimal.Parse(v3);
            }
            if (!string.IsNullOrEmpty(v4))
            {
                d4= decimal.Parse(v4);
            }
            if (!string.IsNullOrEmpty(v5))
            {
                d5 = decimal.Parse(v5);
            }
            if (!string.IsNullOrEmpty(v6))
            {
                d6= decimal.Parse(v6);
            }
            if (etc.CHECK_DATATABLE_IF_EXISTS_DETAIL_COURSE (dt))
            {
                b = true;
            }
            else if (d1 != d2)
            {
                b = true;
                label1.Text = "试算不平衡！年初借方合计不等于年初贷方合计！";

            }
            else if (d3!= d4)
            {
                b = true;
                label1.Text = "试算不平衡！累计借方合计不等于累计贷方合计！";

            }
            else if (d5 != d6)
            {

                b = true;
                label1.Text = "试算不平衡！期初借方合计不等于期初贷方合计！";

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
            
       
        }


        #region dgvDoubleClick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
    
            try
            {
                int currentrowsindex = dataGridView1.CurrentCell.RowIndex;
                int currentcolumnindex = dataGridView1.CurrentCell.ColumnIndex;
                if (currentcolumnindex == 2)
                {
                    CSPSS.BASE_INFO.ACCOUNTANT_COURSE frm = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
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
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }

        }
        #endregion
        #region dgvCellEnter
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
           
               
           
            try
            {
                dgvfoucs();
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
        private void dgvfoucs()
        {

            /*for (i = 0; i < dt.Rows.Count; i++)
            {
                ask(i);
            }*/
            ask(dataGridView1.CurrentCell.RowIndex);

            
        }
        #region ask
        private void ask(int k)
        {

            int n = k;
            decimal 
                v2 = 0,
                v3 = 0,
                v4 = 0,
                v5 = 0;
            dt.Rows[n]["年初借方"] = DBNull.Value;
            dt.Rows[n]["年初贷方"] = DBNull.Value;
            if (!string.IsNullOrEmpty(dt.Rows[k]["累计借方"].ToString()))
            {
                v2 = decimal.Parse(dt.Rows[k]["累计借方"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[k]["累计贷方"].ToString()))
            {
                v3 = decimal.Parse(dt.Rows[k]["累计贷方"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[k]["期初借方"].ToString()))
            {
                v4 = decimal.Parse(dt.Rows[k]["期初借方"].ToString());
            }
            if (!string.IsNullOrEmpty(dt.Rows[k]["期初贷方"].ToString()))
            {
                v5 = decimal.Parse(dt.Rows[k]["期初贷方"].ToString());
            }
            if (v4 > 0)
            {
                dt.Rows[n]["期初借方"] = v4;
            }
            else if (v5 > 0)
            {

                dt.Rows[n]["期初贷方"] = v5;
            }
            decimal
               v7 = v3 - v2 + v4 - v5,
               v8 = v2 - v3 - v4 + v5;
            if (v7 > 0)
            {
                dt.Rows[n]["年初借方"] = v3 - v2 + v4 - v5;
                dt.Rows[n]["年初贷方"] = DBNull.Value;

            }
            else if (v8 > 0)
            {
                dt.Rows[n]["年初贷方"] = v2 - v3 - v4 + v5;
                dt.Rows[n]["年初借方"] = DBNull.Value;
            }

          
            ask1();
        }
        #endregion
        #region ask1
        private void ask1()
        {

          
            t3.Text = "0.00";
            t4.Text = "0.00";
            t5.Text = "0.00";
            t6.Text = "0.00";
            string v1 = dt.Compute("sum(年初借方)", "").ToString();
            string v2 = dt.Compute("sum(年初贷方)", "").ToString();
            string v3 = dt.Compute("sum(累计借方)", "").ToString();
            string v4 = dt.Compute("sum(累计贷方)", "").ToString();
            string v5 = dt.Compute("sum(期初借方)", "").ToString();
            string v6 = dt.Compute("sum(期初贷方)", "").ToString();
      
            if (!string.IsNullOrEmpty(v3))
            {
                t3.Text = string.Format("{0:F2}", Convert.ToDouble(v3));
             
            }
            if (!string.IsNullOrEmpty(v4))
            {
                t4.Text = string.Format("{0:F2}", Convert.ToDouble(v4));
            }
            if (!string.IsNullOrEmpty(v5))
            {
                t5.Text = string.Format("{0:F2}", Convert.ToDouble(v5));

            }
            if (!string.IsNullOrEmpty(v6))
            {
                t6.Text = string.Format("{0:F2}", Convert.ToDouble(v6));
            }
        }
        #endregion
        #region dgvCellValidating
        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 8 && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (e.ColumnIndex == 9 && bc.yesno(e.FormattedValue.ToString()) == 0)
                {
                    e.Cancel = true;
                    MessageBox.Show("只能输入数字！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }

       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #endregion
       
        private void dataGridView1_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (this.dataGridView1.Rows[0].Cells["期初借方"].Value.ToString() != "" && this.dataGridView1.Rows[0].Cells["期初贷方"].Value.ToString() != "")
            {
                e.Cancel = true;
                MessageBox.Show("期初借方与期初贷方同行只能输入一方！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void btnACCOUNTANT_STATUS_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(@"帐套启用后期初开帐数据将不能再修改，您需确认所有录入的数据正确无误再点击，也可暂时放弃此操作，后续再启用！", "提示",
                                             MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                bc.getcom("UPDATE PERIOD SET ACCOUNT_IF_START_USING='Y'");/* aready do initialize*/
             
            }
        }

    

    }
}
