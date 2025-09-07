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

namespace CSPSS
{
    public partial class MAIN : Form
    {
         DataTable dt = new DataTable();
         DataTable dt2 = new DataTable();
         basec bc = new basec();
         CUSER cuser = new CUSER();
         CEMPLOYEE_INFO cemplyee_info = new CEMPLOYEE_INFO();
         Color c2 = System.Drawing.ColorTranslator.FromHtml("#4a7bb8");
         Color c3 = System.Drawing.ColorTranslator.FromHtml("#24ade5");
         CVOUCHER cvoucher = new CVOUCHER();
         CDEPART cdepart = new CDEPART();
         CPOSITION cposition = new CPOSITION();
         CPAYMENT_METHOD cpayment_method = new CPAYMENT_METHOD();
         //CUSER_GROUP cuser_group = new CUSER_GROUP();
        public MAIN()
        {
            InitializeComponent();
        }
        private void MAIN_Load(object sender, EventArgs e)
        {
            this.Text = "财务管理系统 Version 1.0.15";
            dt = bc.getdt("SELECT * from RightList where USID = '"+LOGIN .USID +"'");
            SHOW_TREEVIEW(dt);
            menuStrip1.Font = new Font("宋体", 9);
            this.WindowState = FormWindowState.Maximized;
            toolStripStatusLabel1.Text = "||当前用户：" + LOGIN.UNAME;
            toolStripStatusLabel2.Text = "||所属部门：" + LOGIN.DEPART;
            toolStripStatusLabel3.Text = "||登录时间：" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
            listView1.BackColor = c2;
            listBox3.BackColor = c2;
            listBox3.Height = 84;
            groupBox1.BackColor = c2;
            listView1.ForeColor = Color.White;
            listView1.Font = new Font("新宋体", 11);

            listView1.Location = new Point(1, 75);
            listView2.BorderStyle = BorderStyle.None;
            //listView1.BorderStyle = BorderStyle.None;
            listView2.Location = new Point(195, 75);
            listBox3.Location = new Point(1, -35);
            listView1.Height = 660;
            listView2.Height = 660;
            listView1.Width = 194;
             //listView1 .BackgroundImage  = Image.FromFile(System .IO.Path.GetFullPath("Image/1.png"));
            groupBox1.Height = 675;

            imageList1.Images.Add(CSPSS.Resource1._1);
            imageList1.Images.Add(CSPSS.Resource1._2);
            imageList1.Images.Add(CSPSS.Resource1._3);
            imageList1.Images.Add(CSPSS.Resource1._4);
            imageList1.Images.Add(CSPSS.Resource1._5);
            imageList1.Images.Add(CSPSS.Resource1._6);
            imageList1.Images.Add(CSPSS.Resource1._7);
            imageList1.Images.Add(CSPSS.Resource1._8);
            imageList1.Images.Add(CSPSS.Resource1._9);
            imageList1.Images.Add(CSPSS.Resource1._10);
            imageList1.Images.Add(CSPSS.Resource1._11);
            imageList1.Images.Add(CSPSS.Resource1._12);
            imageList1.Images.Add(CSPSS.Resource1._13);
            imageList1.Images.Add(CSPSS.Resource1._14);
            imageList1.Images.Add(CSPSS.Resource1._15);
            imageList1.Images.Add(CSPSS.Resource1._16);
            imageList1.Images.Add(CSPSS.Resource1._17);
            imageList1.Images.Add(CSPSS.Resource1._18);
            imageList1.Images.Add(CSPSS.Resource1._19);
            imageList1.Images.Add(CSPSS.Resource1._20);
            imageList1.Images.Add(CSPSS.Resource1._21);
            imageList1.Images.Add(CSPSS.Resource1._22);
            imageList1.Images.Add(CSPSS.Resource1._23);
            imageList1.Images.Add(CSPSS.Resource1._24);
            imageList1.Images.Add(CSPSS.Resource1._25);
            imageList1.Images.Add(CSPSS.Resource1._26);
            imageList1.Images.Add(CSPSS.Resource1._27);
            imageList1.Images.Add(CSPSS.Resource1._28);
            imageList1.Images.Add(CSPSS.Resource1._29);
            imageList1.Images.Add(CSPSS.Resource1._30);
            imageList1.Images.Add(CSPSS.Resource1._31);
            imageList1.Images.Add(CSPSS.Resource1._32);
            imageList1.Images.Add(CSPSS.Resource1._33);
            imageList1.Images.Add(CSPSS.Resource1._34);
            imageList1.Images.Add(CSPSS.Resource1._35);
            imageList1.Images.Add(CSPSS.Resource1._36);
            imageList1.Images.Add(CSPSS.Resource1._37);
            imageList1.Images.Add(CSPSS.Resource1._38);
            imageList1.Images.Add(CSPSS.Resource1._39);
            imageList1.Images.Add(CSPSS.Resource1._40);
            imageList1.Images.Add(CSPSS.Resource1._41);
            imageList1.Images.Add(CSPSS.Resource1._42);

            imageList1.ColorDepth = ColorDepth.Depth32Bit;/*防止图片失真*/
            listView1.View = View.SmallIcon;
            listView2.View = View.LargeIcon;
            imageList1.ImageSize = new Size(48, 48);/*set imglist size*/
            listView1.SmallImageList = imageList1;
            listView2.LargeImageList = imageList1;
            bind();
        }
        public void bind()
        {
            PERIOD PE = new PERIOD();
            toolStripStatusLabel4.Text = PE.ACCOUNTING_PERIOD;

        }
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {
           
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "", "PARENT_NODEID=0");

            if (dt.Rows.Count > 0)
            {
                for(int i=0;i<dt.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView1.Items.Add(dt.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
        
               DataTable  dtx = bc.GET_DT_TO_DV_TO_DT(dt, "", "NODE_NAME='账务管理'");
                if (dtx.Rows.Count > 0)
                {
                    click(dtx.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[1].BackColor = c3;
                }
                else
                {

                    click(dt.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[0].BackColor = c3;
                }
               
                
            }
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string NODEID)
        {

            dt2 = bc.getdt("SELECT * FROM RIGHTLIST WHERE PARENT_NODEID='" + NODEID  + "'AND  USID = '" + LOGIN.USID + "'");
            if (dt2.Rows.Count > 0)
            {
                for(int i=0;i<dt2.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView2.Items.Add(dt2.Rows [i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt2.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
            }
        }
        #endregion

         private void 退出系统ToolStripMenuItem1_Click(object sender, EventArgs e)
         {
             if (MessageBox.Show("确定要退出本系统吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
             {
                 Application.Exit();
             }
             else
             {
                 MAIN fmain = new MAIN();
                 fmain.Show();
             }
         }
         private void listView1_Click(object sender, EventArgs e)
         {
            
             string v1 = listView1.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
             click(v1);
            
         }
         private void click(string NODE_NAME)
         {
             listView2.Items.Clear();
             string id = bc.getOnlyString("SELECT NODEID FROM RIGHTLIST WHERE NODE_NAME='" + NODE_NAME + "'");
             SHOW_TREEVIEW_O(id);

             foreach (ListViewItem lvi in listView1.Items)
             {
                 if (lvi.Selected)
                 {
                     lvi.BackColor = c3;
                     pictureBox1.Focus();/*SELECTED AFTER MOVE FOCUS*/
                 }
                 else
                 {
                     lvi.BackColor = c2;
                 }

             }

         }
         private void listView2_Click(object sender, EventArgs e)
         {
             string v1 = listView2.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/

             #region v1
             if (v1 == "科目信息维护")
             {
                 CSPSS.BASE_INFO.ACCOUNTANT_COURSE FRM = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
                 FRM.Show();
                

             }
             else if (v1 == "币别信息维护")
             {
                 try
                 {
                     if (!bc.exists("SELECT * FROM PERIOD"))
                     {
                         MessageBox.Show("未做期初账期作业，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }
                     else
                     {
                         CSPSS.BASE_INFO.CURRENCY FRM = new CSPSS.BASE_INFO.CURRENCY();
                         FRM.ShowDialog();
                     }
                 }
                 catch (Exception)
                 {

                 }
               

             }
            
             else if (v1 == "员工信息维护")
             {
                 CSPSS.BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
                 FRM.IDO = cemplyee_info.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "部门信息维护")
             {
                 CSPSS.BASE_INFO.DEPART FRM = new CSPSS.BASE_INFO.DEPART();
                 FRM.IDO = cdepart.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "职务信息维护")
             {
                 CSPSS.BASE_INFO.POSITION FRM = new CSPSS.BASE_INFO.POSITION();
                 FRM.IDO = cposition.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "期初账期作业")
             {
                 try
                 {
                     if (bc.exists("SELECT * FROM PERIOD"))
                     {
                         if (MessageBox.Show(@"系统已经维护过期初账期，如果重新做期初账期作业，将会删除所有的做账凭证含期初数据及币别信息，且数据不可恢复，请谨慎此操作！", "提示",
                                                         MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                         {
                             bc.getcom("DELETE VOUCHER_DET");/* aready do initialize*/
                             bc.getcom("DELETE VOUCHER_MST");
                             bc.getcom("DELETE PERIOD");
                             //bc.getcom("DELETE CURRENCY_DET");
                             //bc.getcom("DELETE CURRENCY_MST");
                             bind();
                             VOUCHER_MANAGE.INITIAL_PERIOD FRM = new VOUCHER_MANAGE.INITIAL_PERIOD(this);
                             FRM.ShowDialog();
                         }

                     }
                     else
                     {
                         VOUCHER_MANAGE.INITIAL_PERIOD FRM = new VOUCHER_MANAGE.INITIAL_PERIOD(this);
                         FRM.ShowDialog();


                     }
                 }
                 catch (Exception)
                 {

                 }

             }
             else if (v1 == "期初开账作业")
             {
                 if (!bc.exists("SELECT * FROM PERIOD"))
                 {
                     MessageBox.Show("未做期初账期作业，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 else if (bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                 {

                     MessageBox.Show("账套已经启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 else
                 {
                     PERIOD PE = new PERIOD();
                     VOUCHER_MANAGE.INITIAL_CONSULENZA FRM = new VOUCHER_MANAGE.INITIAL_CONSULENZA();
                     FRM.FINANCIAL_YEAR_INITIAL_DATE = PE.FINANCIAL_YEAR_INITIAL_DATE;
                     FRM.EXPIRATION_DATE = PE.EXPIRATION_DATE;
                     FRM.ShowDialog();

                 }
                 try
                 {

                 }
                 catch (Exception)
                 {

                 }

             }
             else if (v1 == "录入凭证作业")
             {
                /* CSPSS.VOUCHER_MANAGE.VOUCHERT FRM = new CSPSS.VOUCHER_MANAGE.VOUCHERT();
                 FRM.IDO = cvoucher.GETID();
                 FRM.ShowDialog();*/
                 string a1 = bc.numYMD(12, 4, "0001", "select * from VOUCHER_MST", "VOID", "VO");
                 if (a1 == "Exceed Limited")
                 {
                     MessageBox.Show("编码超出限制！", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 }
                 else if (!bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                 {
                     MessageBox.Show("账套未启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                 }
                 else
                 {
                     VOUCHER_MANAGE.VOUCHERT FRM = new VOUCHER_MANAGE.VOUCHERT();
                     FRM.ACID = a1;
                     FRM.ShowDialog();
                 }
                 try
                 {

                 }
                 catch (Exception)
                 {

                 }

             }
             else if (v1 == "凭证查询作业")
             {
                /* CSPSS.VOUCHER_MANAGE.VOUCHER FRM = new CSPSS.VOUCHER_MANAGE.VOUCHER();
                 FRM.ShowDialog();*/

                 try
                 {
                     if (!bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                     {
                         MessageBox.Show("账套未启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }
                     else
                     {
                         VOUCHER_MANAGE.VOUCHER FRM = new VOUCHER_MANAGE.VOUCHER();
                         FRM.ShowDialog();
                     }
                 }
                 catch (Exception)
                 {

                 }
            

             }
             else if (v1 == "结转损益作业")
             {
                 try
                 {
                     if (!bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                     {
                         MessageBox.Show("账套未启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }
                     else
                     {
                         VOUCHER_MANAGE.THE_FINAL_BILL FRM = new VOUCHER_MANAGE.THE_FINAL_BILL(this);
                         FRM.ShowDialog();


                     }
                 }
                 catch (Exception)
                 {

                 }

             }
             else if (v1 == "期末结账作业")
             {
                 try
                 {
                     if (!bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                     {
                         MessageBox.Show("账套未启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }
                     else
                     {
                         VOUCHER_MANAGE.CLOSING_ACCOUNT FRM = new VOUCHER_MANAGE.CLOSING_ACCOUNT(this);
                         FRM.ShowDialog();

                     }
                 }
                 catch (Exception)
                 {

                 }

             }
             else if (v1 == "资产负债表")
             {
                 try
                 {
                     if (!bc.exists("SELECT * FROM PERIOD WHERE ACCOUNT_IF_START_USING='Y'"))
                     {
                         MessageBox.Show("账套未启用，此作业不可用！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     }
                     else
                     {
                         REPORT_MANAGE.BALANCE_SEARCH FRM = new REPORT_MANAGE.BALANCE_SEARCH();
                         FRM.ShowDialog();

                     }
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show(ex.Message);
                 }

             }
             else if (v1 == "用户帐户")
             {
                 CSPSS.USER_MANAGE.USER_INFO FRM = new CSPSS.USER_MANAGE.USER_INFO();
                 FRM.IDO = cuser.GETID();
                 FRM.ADD_OR_UPDATE = "ADD";
                 FRM.ShowDialog();
               

             }
             else if (v1 == "更改密码")
             {
                 CSPSS.USER_MANAGE.EDIT_PWD FRM = new CSPSS.USER_MANAGE.EDIT_PWD();
                 FRM.ShowDialog();

             }
             else if (v1 == "权限管理")
             {
                 CSPSS.USER_MANAGE.EDIT_RIGHT FRM = new CSPSS.USER_MANAGE.EDIT_RIGHT();
                 FRM.ShowDialog();

             }
    
             #endregion
         }
    }
}
