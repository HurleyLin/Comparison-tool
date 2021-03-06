﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;
using System.IO;
using Oracle_DLL;
using System.Data.OleDb;
using System.Threading; 

namespace Comparison_tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static public DataTable ExcelToDS(string Path)
        {
            //string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "data source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + Path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataTable dt = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            dt = new DataTable();
            myCommand.Fill(dt);
            return dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime beforDT = System.DateTime.Now;
            int Rows;
            //DataTable dt = new DataTable();

            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();

            if (fd.ShowDialog() == DialogResult.OK)
            {
                string fileName = fd.FileName;
                Rows = 0;
                DataTable p = new DataTable();
                DataTable dt = ExcelToDS(fileName);
                int number = dt.Rows.Count;
                //dataGridView1.DataSource = dt;
                //this.dataGridView1.DataSource = dt;
                if (number > 0)
                {
                    //DataRow dr = null;
                    bool Comparison = false;
                    for (int i = 1; i < number; i++)
                    {
                        //dr = dt.Rows[i];
                        string SN = dt.Rows[i][6].ToString();
                        string CATTONNO = dt.Rows[i][12].ToString();
                        string DATE = dt.Rows[i][14].ToString();

                        if (SN == "")
                            continue;
                        bool result = ORACLEDLL.Comparison(SN,CATTONNO,DATE);
                        if (result)
                        {
                            /*
                            bool res = ORACLEDLL.InsterApicalComparison(SN, CATTONNO, DATE);
                            if (!res)
                            {
                                //MessageBox.Show("录入失败！");
                                textBox2.Text += "SN：" + dt.Rows[i][6].ToString() + "    " + "箱号：" + dt.Rows[i][12].ToString() + "    " + " 录入失败\r\n";
                                Application.DoEvents();
                            }
                            textBox1.Text += "SN：" + dt.Rows[i][6].ToString() + "    " + "箱号：" + dt.Rows[i][12].ToString() + "    " + " 录入成功\r\n"; 
                            Application.DoEvents(); 
                             * */

                            //continue;
                            
                        }
                        else
                        {
                            Comparison = true;
                            textBox1.Text += "SN：" + dt.Rows[i][6].ToString() + "    " + "箱号：" + dt.Rows[i][12].ToString() + "    " + " 重复\r\n";
                            //Application.DoEvents(); 
                        }

                        label4.ForeColor = Color.Red;
                        label4.Text = "正在检测数据，请稍后......";
                        Application.DoEvents();
                        //System.Threading.Thread.Sleep(100);
                    }

                    if (Comparison)
                    {
                        MessageBox.Show("有重复数据，请检查！");
                    }
                    else
                    {

                        label4.ForeColor = Color.Red;
                        label4.Text = "正在录入数据......";
                        for (int i = 1; i < number; i++)
                        {
                            string SN = dt.Rows[i][6].ToString();
                            string CATTONNO = dt.Rows[i][12].ToString();
                            string DATE = dt.Rows[i][14].ToString();

                            if (SN == "")
                                continue;

                            bool res = ORACLEDLL.InsterApicalComparison(SN, CATTONNO, DATE);
                            if (!res)
                            {
                                textBox1.Text += "SN：" + dt.Rows[i][6].ToString() + "    " + "箱号：" + dt.Rows[i][12].ToString() + "    " + " 录入失败\r\n";
                                MessageBox.Show("录入失败！");                        
                                Application.DoEvents();
                            }
                            textBox1.Text += "SN：" + dt.Rows[i][6].ToString() + "    " + "箱号：" + dt.Rows[i][12].ToString() + "    " + " 录入成功\r\n";
                            Rows++;
                            Application.DoEvents();
                        }
                    }

                    string exepath = System.IO.Directory.GetCurrentDirectory();
                    string filepath = exepath + "\\repeat.txt";

                    //FileStream fs = new FileStream("repeat.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    FileStream fs = new FileStream(filepath, FileMode.Append, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs);
                    DateTime.Now.ToString();        //获取当前系统时间 完整的日期和时间
                    sw.WriteLine(DateTime.Now.ToString());
                    sw.WriteLine(textBox1.Text);
                    //sw.WriteLine(textBox2.Text);
                    sw.Flush();//文件流
                    sw.Close();//最后要关闭写入状态
                    if (!Comparison)
                    {
                        label2.Text = "本次操作共录入 " + Rows + "行";
                        MessageBox.Show("导入成功！");
                    }
                        
                }
                else
                {
                    MessageBox.Show("没有数据！");
                }
                label4.ForeColor = Color.Green;
                label4.Text = "操作完成！";

                label2.ForeColor = Color.Red;
            
            }

            DateTime afterDT = System.DateTime.Now;
            TimeSpan ts = afterDT.Subtract(beforDT);
            
        }   
    }
}
