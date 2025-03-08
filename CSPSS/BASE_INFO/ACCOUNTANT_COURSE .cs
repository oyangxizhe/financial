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
    public partial class ACCOUNTANT_COURSE : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        protected int M_int_judge, t;
        basec bc = new basec();
        ExcelToCSHARP etc = new ExcelToCSHARP();
        Color c = System.Drawing.ColorTranslator.FromHtml("#efdaec");
        PERIOD pe = new PERIOD();
        string v;
        private string _ACCODE;
        public string ACCODE
        {
            set { _ACCODE = value; }
            get { return _ACCODE; }
        }
        public ACCOUNTANT_COURSE()
        {
            InitializeComponent();
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
           
        }
        private void ACCOUNTANT_COURSE_Load(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(1);
            bind(dt);
            textBox2.BackColor = Color.Yellow;
            textBox3.BackColor = Color.Yellow;
            try
            {
                
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
            List<string> list1 = etc.getCOURSE_TYPE_INFO();
            for (int i = 0; i < list1.Count; i++)
            {

                comboBox1.Items.Add(list1[i]);
            }
            DataTable dtx1 = bc.getdt("SELECT CYCODE FROM CURRENCY_MST ");
            foreach (DataRow dr in dtx1.Rows)
            {

                comboBox3.Items.Add(dr["CYCODE"].ToString ());
            }
            LoadAgain();
            textBox2.BorderStyle = BorderStyle.FixedSingle;
            textBox3.BorderStyle = BorderStyle.FixedSingle;
            textBox5.BorderStyle = BorderStyle.FixedSingle;
            SHOW_TREEVIEW(dt);
      
            currency();
          
        }
        private void currency()
        {
           
            DataTable dtx = bc.getdt("SELECT * FROM CURRENCY_MST WHERE CYCODE='RMB'");
            if (dtx.Rows.Count > 0)
            {
                comboBox3.Text = dtx.Rows[0]["CYCODE"].ToString();

            }

        }
        #region bind
        private void bind()
        {
            dt = etc.GetCOURSE_LoadData();
            if (dt.Rows.Count > 0)
            {
               

            }
           

            //this.WindowState = FormWindowState.Maximized;
            think();

        }
        #endregion
        #region think
        private void think()
        {

            dt2 = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE ACCODE=SUBSTRING(ACCODE,1,4)");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            AutoCompleteStringCollection inputInfoSource4 = new AutoCompleteStringCollection();
            comboBox2.Items.Clear();
                foreach (DataRow dr in dt2.Rows)
                {

                    comboBox2.Items.Add(dr["ACCODE"].ToString() + " " + dr["ACNAME"].ToString());
                    inputInfoSource.Add(dr["ACCODE"].ToString() + " " + dr["ACNAME"].ToString());


                }
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox2.AutoCompleteCustomSource = inputInfoSource;


        }
        #endregion
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {

           
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "PARENT_NODEID IS NULL");

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    TreeNode trd = treeView1.Nodes.Add(dr["ACCODE"].ToString() + " " + dr["ACNAME"].ToString());

                    if (trd.Text == textBox2.Text + " " + textBox3.Text)
                    {

                        trd.BackColor = c;

                    }
                    SHOW_TREEVIEW_O(dr["ACID"].ToString(), trd);

                }
                MessageBox.Show("ok");
            }
            else
            {


            }
            
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string ACID,TreeNode trd)
        {

                    dt2 = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE PARENT_NODEID='" + ACID + "'");
                    if (dt2.Rows.Count > 0)
                    {
                        foreach (DataRow dr1 in dt2.Rows)
                        {

                            TreeNode TRC = new TreeNode();
                            TRC.Text =dr1["ACCODE"].ToString()+" "+dr1["ACNAME"].ToString ();
                            trd.Nodes.Add(TRC);
                            if (TRC.Text == textBox2.Text+" "+textBox3 .Text )
                            {

                                TRC.BackColor = c;
                                MessageBox.Show("ok");
                            }
                            SHOW_TREEVIEW_O(dr1["ACID"].ToString(),TRC);
                          
                        }
                   }
        }
        #endregion
        #region bind1
        private void bind(DataTable dt)
        {

            try
            {
                if (dt.Rows.Count > 0)
                {
                    textBox1.Text = dt.Rows[0]["ACID"].ToString();
                    textBox2.Text = dt.Rows[0]["ACCODE"].ToString();
                    textBox3.Text = dt.Rows[0]["ACNAME"].ToString();
                    bind2(dt.Rows[0]["ACCODE"].ToString());
                  
                }
                think();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #endregion
        #region bind2
        private void bind2(string ACCODE)
        {
            try
            {
                comboBox1.Text = bc.getOnlyString("SELECT COURSE_TYPE FROM ACCOUNTANT_COURSE WHERE ACCODE='" + ACCODE + "'");
                string v = bc.getOnlyString("SELECT BALANCE_DIRECTION FROM ACCOUNTANT_COURSE WHERE ACCODE='" + ACCODE + "'");
                if (v == "借")
                {
                    radioButton1.Checked = true;
                }
                else
                {
                    radioButton2.Checked = true;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
        #endregion
       
        #region save
        protected void save()
        {
            string v;
            if (radioButton1.Checked == true)
            {
                v = "借";
            }
            else
            {
                v = "贷";
            }
            etc.save(textBox1.Text, textBox2.Text, textBox3.Text, comboBox1.Text, v,comboBox3 .Text );
            
            if (etc.IFExecution_SUCCESS)
            {
                //LoadAgain();
                COURSE_TYPE_LOAD();
               
            }
           
        }
        private void COURSE_TYPE_LOAD()
        {
            if (textBox2.Text.Length > 0)
            {
                int k = Convert.ToInt32(textBox2.Text.Substring(0, 1));
                dt = etc.GetCOURSE_TypeData(k);
            }

            if (dt.Rows.Count > 0)
            {

                //bind(dt);
            }
            else
            {
                textBox1.Text = "";
                ClearText();
               
            }
            think();
            textBox2.Focus();
       
          
            treeView1.Nodes.Clear();// no allow once again onload
            
            SHOW_TREEVIEW(dt);
            if (textBox2.Text.Length >= 4)
            {
                foreach (TreeNode trd in treeView1.Nodes)
                {
                    if (trd.Text.Substring(0, 4) == textBox2.Text.Substring(0, 4))
                    {

                        trd.ExpandAll();
                    

                    }
                    //MessageBox.Show(trd.Text);
                }
            }
            LoadAgain();
            //
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
        
        }

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
        #region excelprint
        private void btnExcelPrint_Click(object sender, EventArgs e)
        {

        }
        #endregion
        #region btnadd

        #endregion
        #region loadagain
        private void LoadAgain()
        {
            ClearText();
            string a1 = bc.numYM(10, 4, "0001", "select * from Accountant_Course", "ACID", "AC");
            if (a1 == "Exceed Limited")
            {
                MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                textBox1.Text = a1;
            }
            //dataGridView1.DataSource = total1();
        }
        #endregion
        private void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            comboBox1.Text = "";
            radioButton1.Checked = false;
            radioButton2.Checked = false;

        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                v = "借";

            }
            else
            {
                v = "贷";

            }
            ExcelToCSHARP etc = new ExcelToCSHARP();
            List<string> list = etc.getCOURSE_TYPE_INFO();
            if (textBox1.Text == "")
            {
                MessageBox.Show("科目编号不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (etc.JuageACCODEFormatt(textBox2.Text, textBox3.Text, comboBox1.Text, v))
            {


            }
            else if (bc.JuageIfAllowKEYIN(list, comboBox1.Text, "科目类别不存在！"))
            {


            }

            else
            {
                if (etc.IfFirstDetailCourse == true && etc.IFCONSULENZA == true)
                {
                    if (MessageBox.Show("如果在科目 " + etc.ACCODE + " 下新增明细科目此科目本年度发生的金额将结转到新增科目" + textBox2.Text +
                        "下，是否继续此操作？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    {
                        basec.getcoms(@"UPDATE VOUCHER_DET SET ACID='" + textBox1.Text +
                            "' FROM VOUCHER_DET A LEFT JOIN VOUCHER_MST B ON A.VOID=B.VOID   WHERE A.ACID='" + etc.ACID +
                            "' AND B.VOUCHER_DATE BETWEEN '" + pe.FINANCIAL_YEAR_STARTDATE + "' AND '" + pe.FINANCIAL_YEAR_ENDDATE + "'");
                        save();

                    }
                }
                else
                {
                    save();

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
                    string v = textBox1.Text;
                    string v1 = bc.getOnlyStringO("ACCOUNTANT_COURSE", "ACNAME", "ACID", v);
                    string v2 = bc.getOnlyStringO("ACCOUNTANT_COURSE", "ACCODE", "ACID", v);

                    if (etc.JuageOnlayOneDetailCourse("ACCOUNTANT_COURSE", "ACCODE", v2, "科目") == 3 && etc.hint != null)
                    {
                        if (MessageBox.Show(etc.hint, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        {
                            basec.getcoms(@"UPDATE VOUCHER_DET SET ACID='" + etc.ACID +
                           "' FROM VOUCHER_DET A LEFT JOIN VOUCHER_MST B ON A.VOID=B.VOID   WHERE A.ACID='" + v +
                           "' AND B.VOUCHER_DATE BETWEEN '" + pe.FINANCIAL_YEAR_STARTDATE + "' AND '" + pe.FINANCIAL_YEAR_ENDDATE + "'");
                            basec.getcoms("DELETE Accountant_Course WHERE ACID='" + v + "'");
                            COURSE_TYPE_LOAD();
                        }

                    }
                    else if (etc.CheckKeyInValueIfExistsDetailCourse("ACCOUNTANT_COURSE", "ACCODE", v2, "科目", " 存在明细科目不允许删除！") != 0)
                    {


                    }
                    else if (bc.exists("VOUCHER_DET", "ACID", v, "科目 " + v1 + " " + "已经有做帐记录不允许删除！"))
                    {

                    }

                    else
                    {
                        basec.getcoms("DELETE Accountant_Course WHERE ACID='" + v + "'");
                        COURSE_TYPE_LOAD();
                    }
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


        }
        #endregion

        private void btnToCSharp_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfv = new OpenFileDialog();
            if (opfv.ShowDialog() == DialogResult.OK)
            {
                /*DataSet ds = new DataSet();
                string tablename = ExcelToCSHARP.GetExcelFirstTableName(opfv .FileName );
                ds = ExcelToCSHARP.importExcelToDataSet(opfv .FileName , tablename);
                DataTable dt = ds.Tables[0];
                dataGridView1.DataSource = dt;*/
                string a = opfv.FileName;
                etc.showdata(a);

                dt = etc.GetCOURSE_TypeData(1);
                bind(dt);
                treeView1.Nodes.Clear();
                SHOW_TREEVIEW(dt);
            }
            try
            {

         
            }
            catch (Exception ex)
            {
                MessageBox.Show("error," + ex.Message);
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(1);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(2);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(3);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(4);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(5);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            dt = etc.GetCOURSE_TypeData(6);
            bind(dt);
            treeView1.Nodes.Clear();
            SHOW_TREEVIEW(dt);
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            /*dt = etc.Search(bc.REMOVE_NAME(comboBox2.Text), textBox5.Text);
            dt = etc.dgvNoShowCourseType(dt);
            if (dt.Rows.Count > 0)
            {
                //treeView1.Nodes.Clear();
                SHOW_TREEVIEW(dt);
               
            }
            else
            {
                MessageBox.Show("没有要查找的相关记录！");

            }
            try
            {
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }*/
           
           /* foreach (TreeNode trd in treeView1.Nodes)
            {

                aws(trd);

            }*/
            dt = bc.getdt("SELECT * FROM ACCOUNTANT_COURSE WHERE ACCODE=SUBSTRING(ACCODE,1,4)");
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "ACCODE LIKE '%"+bc.REMOVE_NAME(comboBox2 .Text )+"%'");
            foreach (DataRow dr in dt.Rows)
            {

                //MessageBox.Show(dr["ACCODE"].ToString());
                treeView1.Nodes.Clear();
                SHOW_TREEVIEW(dt);
            }

            foreach (TreeNode trd in treeView1.Nodes)
            {
              

                    trd.ExpandAll();


                
                //MessageBox.Show(trd.Text);
            }

        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
           
        }
     private void aws(TreeNode trd )
     {
         //MessageBox.Show(trd.Text);
         if (trd.Text ==comboBox2 .Text )
         {
             trd.BackColor = c;
             trd.Checked = true;

          
         }
         foreach (TreeNode trd1 in trd.Nodes)
         {
          
             //MessageBox.Show(trd1.Text);
             aws(trd1);

         }



      }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {

            }
            else if (textBox2.Text.Length > 4)
            {

                bind2(textBox2.Text.Substring(0, 4));


            }

        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            
        }
        public void a5()
        {
            
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            LoadAgain();
            textBox2.Focus();
            currency();
        
        }

        private void treeView1_Click(object sender, EventArgs e)
        {
          
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode trd = treeView1.SelectedNode;
            //MessageBox.Show(trd.Index.ToString() + "-" + trd.Text);
            textBox2.Text = bc.REMOVE_NAME(trd.Text);
            textBox3.Text = bc.getOnlyString("SELECT ACNAME FROM ACCOUNTANT_COURSE WHERE ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            textBox1.Text = bc.getOnlyString("SELECT ACID FROM ACCOUNTANT_COURSE WHERE ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            comboBox3.Text = bc.getOnlyString(@"
SELECT B.CYCODE FROM ACCOUNTANT_COURSE  A LEFT JOIN CURRENCY_MST B ON A.CYID=B.CYID WHERE A.ACCODE='" + bc.REMOVE_NAME(trd.Text) + "'");
            bind2(bc.REMOVE_NAME(trd.Text));
            if (trd.IsExpanded)
            {
              
                    trd.Collapse();
                
            }
            else
            {
                trd.Expand();

            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
       
            ACCODE = textBox2.Text;
            this.Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ACCODE = textBox2.Text;
            this.Close();
        }
    }
}
