using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;



namespace 题录数据转换
{
    public partial class frm_txttrans : Form
    {
        public frm_txttrans()
        {
            InitializeComponent();
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
            
        }

        private void btn_openfile_Click(object sender, EventArgs e)
        {
            ofd_findfile.ShowDialog();
            tbx_filedir.Text = ofd_findfile.FileName;
        }

        private void btn_transfer_Click(object sender, EventArgs e)
        {
            string openfile = "";
            openfile = tbx_filedir.Text.Trim();
            if (openfile == "")
            {
                MessageBox.Show("请选择需要转换的文件！","系统提示",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                return;
            }
            if (System.IO.File.Exists (openfile)==false )
            {
                MessageBox.Show("选择的文件不存在！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            string dbconnstr = ConfigurationManager.ConnectionStrings["dbconnString"].ToString();
            string tablename = ConfigurationManager.AppSettings["tablename"].ToString();
            string sql_str = "";
            int j = int.Parse(lbl_datacountshow.Text);
            
            if (radiobtn_cnki.Checked)
            {//CNKI格式选中
                /* Title-题名:
                 * Author-作者:
                 * Organ-单位:
                 * Source-文献来源:
                 * Keyword-关键词:
                 * Summary-摘要:
                 * PubTime-发表时间:
                 * Year-年:
                 * Fund-基金:
                 */
                string title = "",title_tmp="", author = "",author_tmp="", unit = "", source = "", keywords = "", pubtime = "", pubyear = "", fundname = "";
                string dataresource = "CNKI";
                string head_title="Title-题名:";
                string head_author="Author-作者:";
                string head_unit="Organ-单位:";
                string head_source="Source-文献来源:";
                string head_keywords="Keyword-关键词:";
                string head_pubtime="PubTime-发表时间:";
                string head_year="Year-年:";
                string head_fund="Fund-基金:";
                int i=0;

                OleDbConnection strConnection = new OleDbConnection(dbconnstr);
                strConnection.Open();

               // FileStream txtreader= new FileStream(openfile,Encoding.Default );
                StreamReader txtread = new StreamReader(openfile, Encoding.UTF8);
                string rline = txtread.ReadLine();
                do
                {
                    string templine = rline.ToString();
                    if (templine.Length > 4)
                    { 
                        string casekey=templine.Substring(0, templine.IndexOf(":")+1);
                        switch(casekey)
                        {
                            case "Title-题名:":
                                if (title_tmp == "")
                                {  
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;
                                }
                                else
                                { //执行写入数据操作，写入输入ACCESS库

                                    sql_str = "insert into " + tablename + " (title,author,unit,source,keywords,pubtime,fund,pubyear,dataresource) values (";
                                    sql_str += "'"+title+"',";
                                    sql_str += "'"+author+"',";
                                    sql_str += "'"+unit+"',";
                                    sql_str += "'"+source+"',";
                                    sql_str += "'"+keywords+"',";
                                    sql_str += "'"+pubtime+"',";
                                    sql_str += "'"+fundname+"',";
                                    sql_str += "'"+pubyear+"',";
                                    sql_str += "'" + dataresource + "')";
                                    //MessageBox.Show(sql_str);

                                    //数据库写入
                                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                                    try
                                    {
                                        dm.ExecuteNonQuery();
                                        i += 1; j += 1;
                                        lbl_transcountshow.Text = "CNKI题录" + i.ToString() + "条记录";
                                        lbl_datacountshow.Text = j.ToString();
                                        this.Refresh();
                                        title = author = unit =source = keywords = pubtime = pubyear = fundname = "";

                                    }
                                    catch (System.Exception  test)
                                    {
                                        MessageBox.Show(test.ToString(),test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    finally
                                    {
                                        dm.Dispose();
                                        sql_str = "";
                                        //title = "";
                                    }
                                    //开始新一轮读取
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;
                                    // OleDbDataAdapter data_adapter = new OleDbDataAdapter(sql_str, strConnection);
                                }                       
                                break;
                            case "Author-作者:":
                                author = templine.Substring(head_author.Length).Trim();
                                author_tmp = author;
                                break;
                            case "Organ-单位:":
                                unit = templine.Substring(head_unit.Length).Trim();
                                break;
                            case "Source-文献来源:":
                                source = templine.Substring(head_source.Length).Trim();
                                break;
                            case "Keyword-关键词:":
                                keywords = templine.Substring(head_keywords.Length).Trim();
                                break;
                            case "PubTime-发表时间:":
                                pubtime = templine.Substring(head_pubtime.Length).Trim();
                                break;
                            case "Year-年:":
                                pubyear = templine.Substring(head_year.Length).Trim();
                                break;
                            case "Fund-基金:":
                                fundname = templine.Substring(head_fund.Length).Trim();
                                break;
                            default :
                                if (title_tmp == "" || author_tmp == "")
                                {
                                    MessageBox.Show("请确认选择的是否是CNKI导出的数据格式！！！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                                break;              
                        }
                    }
                    
                } while ((rline = txtread.ReadLine()) != null);
                //写入最后一条记录
                if (title_tmp != "")
                {
                    sql_str = "insert into " + tablename + " (title,author,unit,source,keywords,pubtime,fund,pubyear,dataresource) values (";
                    sql_str += "'" + title + "',";
                    sql_str += "'" + author + "',";
                    sql_str += "'" + unit + "',";
                    sql_str += "'" + source + "',";
                    sql_str += "'" + keywords + "',";
                    sql_str += "'" + pubtime + "',";
                    sql_str += "'" + fundname + "',";
                    sql_str += "'" + pubyear + "',";
                    sql_str += "'" + dataresource + "')";
                    //MessageBox.Show(sql_str);
                    //数据库写入
                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                    try
                    {
                        dm.ExecuteNonQuery();
                        i += 1; j += 1;
                        lbl_transcountshow.Text = "CNKI题录" + i.ToString() + "条记录";
                        lbl_datacountshow.Text = j.ToString();
                        this.Refresh();
                        title = author = unit = source = keywords = pubtime = pubyear = fundname = "";

                    }
                    catch (System.Exception test)
                    {
                        MessageBox.Show(test.ToString(), test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally
                    {
                        dm.Dispose();

                    }
                }
                MessageBox.Show("转换完毕！");
                strConnection.Close();
            }
            else if (radiobtn_wangfang.Checked)
            {//万方格式选中
                /* 【篇名】
                 * 【作者】
                 * 【学位类型】
                 * 【授予单位】     【作者单位】
                 * 【导师】
                 * 【年份】
                 * 
                 * 【会议名称】
                 * 【出处】
                 * 【年份】
                 */
                string title = "", title_tmp = "", author = "", author_tmp = "", unit1 = "",unit2="",degree="",tutor="", source = "", conferencename = "", pubyear = "";
                string dataresource = "万方数据";
                string degreethesis = "学位论文";
                string conferencepaper = "会议论文";
                //通用头部
                string head_title = "【篇名】";
                string head_author = "【作者】";
                string head_pubyear = "【年份】";
                //学位论文头部
                string head_degree = "【学位类型】";
                string head_unit1 = "【授予单位】";
                string head_tutor = "【导师】";
                //会议论文头部
                string head_unit2 = "【作者单位】";
                string head_conferencename = "【会议名称】";
                string head_source = "【出处】";
                int i = 0;

                OleDbConnection strConnection = new OleDbConnection(dbconnstr);
                strConnection.Open();
                StreamReader txtread = new StreamReader(openfile, Encoding.UTF8);
                string rline = txtread.ReadLine();
                do  //循环体开始
                {
                    string templine = rline.ToString();
                    if (templine.Length > 6)
                    {
                        string casekey = templine.Substring(0, templine.IndexOf("】") + 1);
                        switch (casekey)
                        {
                            case "【篇名】":
                                if (title_tmp == "")
                                {
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;//MessageBox.Show(title);
                                }
                                else
                                {
                                    //判断是会议论文还是学术论文
                                    if (degree != "" || unit1 != "" || tutor != "")
                                    {
                                        //执行写入数据操作，写入输入ACCESS库
                                        sql_str = "insert into " + tablename + " (title,author,unit,degree,tutor,pubyear,papertype,dataresource) values (";
                                        sql_str += "'" + title + "',";
                                        sql_str += "'" + author + "',";
                                        sql_str += "'" + unit1 + "',";
                                        sql_str += "'" + degree + "',";
                                        sql_str += "'" + tutor + "',";
                                        sql_str += "'" + pubyear + "',";
                                        sql_str += "'" + degreethesis + "',";
                                        sql_str += "'" + dataresource + "')"; 
                                    }
                                    else
                                    {
                                        //执行写入数据操作，写入输入ACCESS库
                                        sql_str = "insert into " + tablename + " (title,author,unit,source,conferencename,pubyear,papertype,dataresource) values (";
                                        sql_str += "'" + title + "',";
                                        sql_str += "'" + author + "',";
                                        sql_str += "'" + unit2 + "',";
                                        sql_str += "'" + source + "',";
                                        sql_str += "'" + conferencename + "',";
                                        sql_str += "'" + pubyear + "',";
                                        sql_str += "'" + conferencepaper + "',";
                                        sql_str += "'" + dataresource + "')";

                                    }
                                    //MessageBox.Show(sql_str);
                                    //数据库写入
                                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                                    try
                                    {
                                        dm.ExecuteNonQuery();
                                        i += 1; j += 1;
                                        lbl_transcountshow.Text = "万方题录" + i.ToString() + "条记录";
                                        lbl_datacountshow.Text = j.ToString();
                                        this.Refresh();

                                    }
                                    catch (System.Exception test)
                                    {
                                        MessageBox.Show(test.ToString(), test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    finally
                                    {
                                        dm.Dispose();
                                        sql_str = "";
                                        title = title_tmp = author = author_tmp = unit1 = unit2 = degree = tutor = source = conferencename = pubyear = "";
                                    }

                                    //开始新一轮读取
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;
                                //    // OleDbDataAdapter data_adapter = new OleDbDataAdapter(sql_str, strConnection);
                                    
                                }
                                break;
                            case "【作者】":
                                author = templine.Substring(head_author.Length).Trim();
                                author_tmp = author;
                                break;
                            case "【年份】":
                                pubyear = templine.Substring(head_pubyear.Length).Trim();
                                //MessageBox.Show(pubyear.Contains(".").ToString());
                                if (pubyear.Contains("."))
                                {
                                    pubyear=pubyear.Substring(0,pubyear.LastIndexOf("."));

                                }
                                break;
                            // ////////////////////////////
                            //如何区分学位论文和会议论文//
                            // //////////////////////////
                            case "【学位类型】":
                                degree = templine.Substring(head_degree.Length).Trim();
                                break;
                            case "【授予单位】":
                                unit1 = templine.Substring(head_unit1.Length).Trim();
                                break;
                            case "【导师】":
                                tutor = templine.Substring(head_tutor.Length).Trim();
                                break;
                            case "【作者单位】":
                                unit2 = templine.Substring(head_unit2.Length).Trim();
                                break;
                            case "【会议名称】":
                                conferencename = templine.Substring(head_conferencename.Length).Trim();
                                break;
                            case "【出处】":
                                source = templine.Substring(head_source.Length).Trim();
                                break;


                            default :
                                if (title_tmp == "" || author_tmp == "")
                                {
                                    MessageBox.Show("请确认选择的是否是万方数据库导出的数据格式！！！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                                break;
                        
                        }
                    }

                } while ((rline = txtread.ReadLine()) != null);
                //循环体结束

                //写入最后一条记录
                if (title_tmp != "")
                {
                    //判断是会议论文还是学术论文
                    if (degree != "" || unit1 != "" || tutor != "")
                    {
                        //执行写入数据操作，写入输入ACCESS库
                        sql_str = "insert into " + tablename + " (title,author,unit,degree,tutor,pubyear,papertype,dataresource) values (";
                        sql_str += "'" + title + "',";
                        sql_str += "'" + author + "',";
                        sql_str += "'" + unit1 + "',";
                        sql_str += "'" + degree + "',";
                        sql_str += "'" + tutor + "',";
                        sql_str += "'" + pubyear + "',";
                        sql_str += "'" + degreethesis + "',";
                        sql_str += "'" + dataresource + "')";
                    }
                    else
                    {
                        //执行写入数据操作，写入输入ACCESS库
                        sql_str = "insert into " + tablename + " (title,author,unit,source,conferencename,pubyear,papertype,dataresource) values (";
                        sql_str += "'" + title + "',";
                        sql_str += "'" + author + "',";
                        sql_str += "'" + unit2 + "',";
                        sql_str += "'" + source + "',";
                        sql_str += "'" + conferencename + "',";
                        sql_str += "'" + pubyear + "',";
                        sql_str += "'" + conferencepaper + "',";
                        sql_str += "'" + dataresource + "')";

                    }
                    //MessageBox.Show(sql_str);
                    //数据库写入
                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                    try
                    {
                        dm.ExecuteNonQuery();
                        i += 1; j += 1;
                        lbl_transcountshow.Text = "万方题录" + i.ToString() + "条记录";
                        lbl_datacountshow.Text = j.ToString();
                        this.Refresh();

                    }
                    catch (System.Exception test)
                    {
                        MessageBox.Show(test.ToString(), test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally
                    {
                        dm.Dispose();
                    }
                }
                MessageBox.Show("转换完毕！");
                strConnection.Close();
            }
            else if (radiobtn_vip.Checked)
            {//维普格式选中
                /* 【题　名】
                 * 【作　者】
                 * 【机　构】
                 * 【刊　名】
                 * 【关键词】
                 */
                string title="",title_tmp="",author="",author_tmp="", unit="", source="", keywords="", pubyear="";
                string dataresource = "中文维普";
                string head_title = "【题　名】";
                string head_author = "【作　者】";
                string head_unit = "【机　构】";
                string head_source = "【刊　名】";  //情报学报.2010(5).-889-896
                string head_keywords = "【关键词】";
                int i = 0;

                OleDbConnection strConnection = new OleDbConnection(dbconnstr);
                strConnection.Open();

                StreamReader txtread = new StreamReader(openfile, Encoding.GetEncoding("GB2312"));
                string rline = txtread.ReadLine();

                do
                {
                    string templine = rline.ToString();
                    if (templine.Length > 6)
                    {
                        string casekey = templine.Substring(0, templine.IndexOf("】") + 1);
                        switch (casekey)
                        {
                            case "【题　名】":
                                if (title_tmp == "")
                                {
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;//MessageBox.Show(title);
                                }
                                else 
                                {
                                    //执行写入数据操作，写入输入ACCESS库
                                    sql_str = "insert into " + tablename + " (title,author,unit,source,keywords,pubyear,dataresource) values (";
                                    sql_str += "'" + title + "',";
                                    sql_str += "'" + author + "',";
                                    sql_str += "'" + unit + "',";
                                    sql_str += "'" + source + "',";
                                    sql_str += "'" + keywords + "',";
                                    sql_str += "'" + pubyear + "',";
                                    sql_str += "'" + dataresource + "')";
                                    //MessageBox.Show(sql_str);

                                    //数据库写入
                                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                                    try
                                    {
                                        dm.ExecuteNonQuery();
                                        i += 1; j += 1;
                                        lbl_transcountshow.Text = "维普题录" + i.ToString() + "条记录";
                                        lbl_datacountshow.Text = j.ToString();
                                        this.Refresh();

                                    }
                                    catch (System.Exception test)
                                    {
                                        MessageBox.Show(test.ToString(), test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                    finally
                                    {
                                        dm.Dispose();
                                        sql_str = "";
                                        //title = "";
                                        title = author = unit =source = keywords =  pubyear = "";
                                    }
                                    //开始新一轮读取
                                    title = templine.Substring(head_title.Length).Trim();
                                    title_tmp = title;
                                    // OleDbDataAdapter data_adapter = new OleDbDataAdapter(sql_str, strConnection);
                                }
                                break;
                            case "【作　者】":
                                author = templine.Substring(head_author.Length).Trim();
                                author_tmp = author;
                                break;
                            case "【机　构】":
                                unit = templine.Substring(head_unit.Length).Trim();
                                break;
                            case "【刊　名】":
                                //刊名字段包含 年代和巻期。
                                string strtmp = templine.Substring(head_source.Length).Trim();
                                string[] splitstr = strtmp.Split(new char[] { '.' });
                                source = splitstr[0];
                                pubyear = splitstr[1];
                                break;
                            case "【关键词】":
                                keywords = templine.Substring(head_keywords.Length).Trim();
                                break;
                            default :
                                if (title_tmp == "" || author_tmp == "")
                                {
                                    MessageBox.Show("请确认选择的是否是维普数据库导出的数据格式！！！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }
                                break;
                        }
                     }

                } while ((rline = txtread.ReadLine()) != null);

                if (title_tmp !="")
                {
                    //执行写入数据操作，写入输入ACCESS库
                    sql_str = "insert into " + tablename + " (title,author,unit,source,keywords,pubyear,dataresource) values (";
                    sql_str += "'" + title + "',";
                    sql_str += "'" + author + "',";
                    sql_str += "'" + unit + "',";
                    sql_str += "'" + source + "',";
                    sql_str += "'" + keywords + "',";
                    sql_str += "'" + pubyear + "',";
                    sql_str += "'" + dataresource + "')";
                    //MessageBox.Show(sql_str);

                    //数据库写入
                    OleDbCommand dm = new OleDbCommand(sql_str, strConnection);
                    try
                    {
                        dm.ExecuteNonQuery();
                        i += 1; j += 1;
                        lbl_transcountshow.Text = "维普题录" + i.ToString() + "条记录";
                        lbl_datacountshow.Text = j.ToString();
                        this.Refresh();

                    }
                    catch (System.Exception test)
                    {
                        MessageBox.Show(test.ToString(), test.Message.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    finally
                    {
                        dm.Dispose();
                    }
                }
                MessageBox.Show("转换完毕！");
                strConnection.Close();
            }
            else
            {
                MessageBox.Show("请选择题录数据类型！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void btn_toexcel_Click(object sender, EventArgs e)
        {
            string dbconn_str = ConfigurationManager.ConnectionStrings["dbconnString"].ToString();
            string tablename = ConfigurationManager.AppSettings["tablename"].ToString();
            string sql_str = "";
            OleDbConnection dbconn = new OleDbConnection(dbconn_str);
            dbconn.Open();
            sql_str = "select * from " + tablename + " order by id";

            //创建数据连接器
            OleDbDataAdapter da_read=new OleDbDataAdapter(sql_str,dbconn);
            DataSet ds = new DataSet();//创建空数据集
            da_read.Fill(ds, "题录数据");//使用数据连接器填充数据集
            
            //创建EXCEL文件对象
            excel.Application excelapp = new excel.Application();
            excelapp.Application.Workbooks.Add(true);
            excel.Workbook newbook=(excel.Workbook)excelapp.ActiveWorkbook;
            excel.Worksheet newsheet=(excel.Worksheet)newbook.ActiveSheet;
            //写入字段名
            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
            {
                newsheet.Cells[1, i + 1] = ds.Tables[0].Columns[i].ColumnName.ToString();
            }

            //写入数值
            for (int r = 0; r < ds.Tables[0].Rows.Count; r++)
            {
                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                {
                    newsheet.Cells[r + 2, i + 1] = ds.Tables[0].Rows[r][i];
                }
                System.Windows.Forms.Application.DoEvents();
            }
            newsheet.Columns.EntireColumn.AutoFit();
            //选择保存文件
            if (sfd_saveexcel.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                string savefilename = sfd_saveexcel.FileName;
                newsheet.SaveAs(savefilename);
                excelapp.Visible = true;
            }
            dbconn.Close();
            ds.Dispose();
            da_read.Dispose();

        }

        private void frm_txttrans_Load(object sender, EventArgs e)
        {
            /* list表的记录显示 */
            string dbconn_str = ConfigurationManager.ConnectionStrings["dbconnString"].ToString();
            string tablename = ConfigurationManager.AppSettings["tablename"].ToString();
            string sql_str =  "select count(*) as countid from " + tablename;
            OleDbConnection dbconn = new OleDbConnection(dbconn_str);
            dbconn.Open();
            OleDbCommand sql_exe=new OleDbCommand(sql_str,dbconn);
            OleDbDataReader  dbreader=sql_exe.ExecuteReader();
            dbreader.Read();
            lbl_datacountshow.Text=dbreader["countid"].ToString();

            /* medline 和CBM表的记录显示   */

            string medlinetable = ConfigurationManager.AppSettings["medlinetable"].ToString();
            string cbmtable = ConfigurationManager.AppSettings["CBMtable"].ToString();
            sql_str = "select count(*) as countid from " + medlinetable;
            sql_exe = new OleDbCommand(sql_str, dbconn);
            dbreader = sql_exe.ExecuteReader();
            dbreader.Read();
            lbl_datacountshow2.Text = dbreader["countid"].ToString();

            sql_str = "select count(*) as countid from " + cbmtable;
            sql_exe = new OleDbCommand(sql_str, dbconn);
            dbreader = sql_exe.ExecuteReader();
            dbreader.Read();
            lbl_datacountshow3.Text = dbreader["countid"].ToString();


            dbreader.Close();
            dbconn.Close();
            sql_exe.Dispose();
            dbreader.Dispose();
            dbconn.Dispose();
        }

        private void btn_cleandata_Click(object sender, EventArgs e)
        {
            string dbconn_str = ConfigurationManager.ConnectionStrings["dbconnString"].ToString();
            string tablename = ConfigurationManager.AppSettings["tablename"].ToString();
            string sql_str = "delete from " + tablename + " where id >0";

            OleDbConnection dbconn = new OleDbConnection(dbconn_str);
            dbconn.Open();
            if (MessageBox.Show("确认删除所有记录么？", "系统提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                OleDbCommand sql_exe = new OleDbCommand(sql_str, dbconn);
                sql_exe.ExecuteNonQuery();
                MessageBox.Show("数据库已经清空！！！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);




                sql_str = "select count(*) as countid from " + tablename;
                sql_exe = new OleDbCommand(sql_str, dbconn);
                OleDbDataReader dbreader = sql_exe.ExecuteReader();
                dbreader.Read();
                lbl_datacountshow.Text = dbreader["countid"].ToString();
                dbreader.Close();
                sql_exe.Dispose();
                dbreader.Dispose();
            }

            dbconn.Close();
            dbconn.Dispose();

        }

        private void btn_openfile2_Click(object sender, EventArgs e)
        {
            ofd_findfile.ShowDialog();
            tbx_filedir_mesh.Text = ofd_findfile.FileName;
        }

        private void btn_transMESH_Click(object sender, EventArgs e)
        {
            string openfile = "";
            openfile = tbx_filedir_mesh.Text.Trim();
            if (openfile == "")
            {
                MessageBox.Show("请选择需要转换的文件！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (System.IO.File.Exists(openfile) == false)
            {
                MessageBox.Show("选择的文件不存在！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string dbconnstr = ConfigurationManager.ConnectionStrings["dbconnString"].ToString();
            string sql_str = "";
            int j = int.Parse(lbl_datacountshow.Text);

            if (radiobtn_medline.Checked)
            {
                string tablename = ConfigurationManager.AppSettings["medlinetable"].ToString();
                //MEDLINE格式选中
                /*TI:    
                 *AU:    
                 *SO:    
                 *LA:   
                 *AB:     
                 *MJME:  
                 *MIME:  
                 *TG:   
                 *SH:    
                 */
                string title = "", author = "", source = "",lang="", mjme = "", mime="", tgword = "", shword = "";
                string head_title = "TI:";
                string head_author = "AU:";
                string head_source = "SO:";
                string head_LANG = "LA:";
                string head_mjme = "MJME:";
                string head_mime = "MIME:";
                string head_TG = "TG:";
                string head_SH = "SH:";
                int i = 0;






            }
            else if (radiobtn_CBM.Checked)
            {
                string tablename = ConfigurationManager.AppSettings["medlinetable"].ToString();

            }
            else
            {
                MessageBox.Show("请选择题录数据类型！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }




        }


    }
}
 
