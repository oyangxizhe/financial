using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using XizheC;

namespace CSPSS.VOUCHER_MANAGE
{
    public partial class VOUCHER : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dtx = new DataTable();
        basec bc = new basec();
      
        

        protected int M_int_judge, i, look;
        protected int getdata;
        CVOUCHER vou = new CVOUCHER();
      
        public VOUCHER()
        {
            InitializeComponent();
        }
        #region init
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(VOUCHER));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnToExcel = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.PictureBox();
            this.btnExit = new System.Windows.Forms.PictureBox();
            this.btnSearch = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Location = new System.Drawing.Point(0, 216);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(943, 400);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            this.dataGridView1.DataSourceChanged += new System.EventHandler(this.dataGridView1_DataSourceChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btnToExcel);
            this.groupBox1.Location = new System.Drawing.Point(3, 127);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(936, 84);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "查询条件";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // btnToExcel
            // 
            this.btnToExcel.FlatAppearance.BorderSize = 0;
            this.btnToExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnToExcel.Font = new System.Drawing.Font("宋体", 9F);
            this.btnToExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnToExcel.Image")));
            this.btnToExcel.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnToExcel.Location = new System.Drawing.Point(841, 13);
            this.btnToExcel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnToExcel.Name = "btnToExcel";
            this.btnToExcel.Size = new System.Drawing.Size(50, 64);
            this.btnToExcel.TabIndex = 11;
            this.btnToExcel.Text = "导出";
            this.btnToExcel.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.btnToExcel.UseVisualStyleBackColor = false;
            this.btnToExcel.Click += new System.EventHandler(this.btnToExcel_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(857, 95);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(29, 12);
            this.label11.TabIndex = 29;
            this.label11.Text = "退出";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(771, 95);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(29, 12);
            this.label12.TabIndex = 28;
            this.label12.Text = "搜索";
            this.label12.Visible = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.btnAdd);
            this.groupBox2.Controls.Add(this.btnExit);
            this.groupBox2.Controls.Add(this.btnSearch);
            this.groupBox2.Location = new System.Drawing.Point(3, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(936, 121);
            this.groupBox2.TabIndex = 34;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "菜单栏";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(28, 95);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(29, 12);
            this.label17.TabIndex = 24;
            this.label17.Text = "新增";
            this.label17.Click += new System.EventHandler(this.label17_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Image = ((System.Drawing.Image)(resources.GetObject("btnAdd.Image")));
            this.btnAdd.InitialImage = null;
            this.btnAdd.Location = new System.Drawing.Point(12, 20);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(60, 60);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.TabStop = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnExit
            // 
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.InitialImage = null;
            this.btnExit.Location = new System.Drawing.Point(843, 20);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(60, 60);
            this.btnExit.TabIndex = 19;
            this.btnExit.TabStop = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnSearch
            // 
            this.btnSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnSearch.Image")));
            this.btnSearch.InitialImage = null;
            this.btnSearch.Location = new System.Drawing.Point(757, 20);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(60, 60);
            this.btnSearch.TabIndex = 18;
            this.btnSearch.TabStop = false;
            this.btnSearch.Visible = false;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // VOUCHER
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(245)))), ((int)(((byte)(255)))));
            this.ClientSize = new System.Drawing.Size(942, 616);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "VOUCHER";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "凭证查询作业";
            this.Load += new System.EventHandler(this.VOUCHER_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnAdd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnExit)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnSearch)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion
  
        #region Bind
        public  void Bind()
        {

            try
            {
                this.WindowState = FormWindowState.Maximized;
                dt = basec.getdts(vou.getsqlX + " WHERE B.STATUS!='INITIAL' ORDER BY A.VOKEY");
                dataGridView1.DataSource = dt;
                dt1 = bc.getdt("SELECT VOID FROM VOUCHER_MST");
                AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
                dgvStateControl();
            }
            catch (Exception)
            {


            }
        }
        #endregion
        private void VOUCHER_Load(object sender, EventArgs e)
        {

            Bind();
         
        }
        #region dgvStateControl
        private void dgvStateControl()
        {
            dataGridView1.Columns["单价"].Width = 80;
            dataGridView1.Columns["数量"].Width = 80;
            dataGridView1.Columns["币别"].Width = 40;
            dataGridView1.Columns["项次"].Width = 40;
            dataGridView1.Columns["科目代码"].Width = 100;
            dataGridView1.Columns["科目名称"].Width = 120;
            //dataGridView1.Columns["日期"].Width = 80;
            dataGridView1.Columns["凭证号"].Width = 100;
            dataGridView1.Columns["摘要"].Width = 200;
            dataGridView1.Columns["借方原币金额"].Width = 100;
            dataGridView1.Columns["借方本币金额"].Width = 100;
            //dataGridView1.Columns["方向"].Width = 40;
            dataGridView1.Columns["汇率"].Width = 60;
             dataGridView1.Columns["贷方原币金额"].Width = 100;
            dataGridView1.Columns["贷方本币金额"].Width = 100;
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
    

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

        #region add

        #endregion
        #region look

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


        #region doubleclick
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

            if (getdata != 0)
            {
                if (getdata == 1)
                {
                    int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
                    string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[0].Value.ToString().Trim();
                    string s2 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[18].Value.ToString().Trim();
                    string s3 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[19].Value.ToString().Trim();
                    string s4 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[20].Value.ToString().Trim();
                    string s5 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[21].Value.ToString().Trim();
                    /*CSPSS.VOUCHER_MANAGE.FrmSellTableT.data4[0] = "doubleclick";
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[0] = s1;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[1] = s2;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[2] = s3;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[3] = s4;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[4] = s5;*/

                    this.Close();
                }

            }
            else
            {
                VOUCHERT frm = new VOUCHERT(this);
                frm.ACID = dt.Rows[dataGridView1.CurrentCell.RowIndex]["凭证号"].ToString();
                frm.ShowDialog();
            }
        }
        #endregion
        public void a2()
        {

            getdata = 1;

        }
        public void a3()
        {

            getdata = 2;

        }
   
        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
               
                bc.dgvtoExcel(dataGridView1, "凭证明细");
                
            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string a1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_MST", "VOID", "VO");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {

                //VOUCHERT.str1[0] = a1;
                VOUCHERT frm = new VOUCHERT(this);
                frm.ACID = a1;
                frm.ShowDialog();

            }
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                //ak();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }

        }
        #region ak()
        private void ak()
        {

    
        }
        #endregion
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            PERIOD period = new PERIOD();
            MessageBox.Show(period.NEXT_FINANCIAL_YEAR + "," + period.NEXT_PERIOD+","+period .NEXT_PERIOD_t );
        }

  
    }
}
