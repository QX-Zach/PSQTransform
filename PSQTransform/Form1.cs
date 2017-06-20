using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace PSQTransform
{
    public partial class Form1 : Form
    {
        //D:\\personal\\test.xlsx

        private static string other = "其他";
        private string filepath = @"D:\pqs.xlsx";
        private Questionnaire question;
      
        private object Nothing = Missing.Value;

        Excel._Application mApp;
        Excel._Workbook myBook;
        Sheets sheets;
        private int rowCount = 1;
        _Worksheet mySheet;

        private int currentTab = 1;

        public Form1()
        {
            InitializeComponent();
            init();
        }

        private void init()
        {
           
            question = new Questionnaire();
            question.school = 1;
            question.area = 1;
            question.grade = 4;

            ques1_comBox.SelectedIndex = 0;
            ques2_checkListBox.ClearSelected();
            ques2_more_text.Clear();
            ques3_comBox.SelectedIndex = 0;
            ques3_text_more.Clear();
        
            ques4_comBox.SelectedIndex = 0;
            ques5_comBox.SelectedIndex = 0;
            ques6_comBox.SelectedIndex = 0;
            ques7_comBox.SelectedIndex = 0;
            ques8_comBox.SelectedIndex = 0;
            ques8_text_more.Clear();
            ques9_comBox.SelectedIndex = 0;
            ques10_comBox.SelectedIndex = 0;
            ques11_comBox.SelectedIndex = 0;
            ques12_checkListBox.ClearSelected();
            ques12_text_more.Clear();

            ques13_comBox.SelectedIndex = 0;
            ques14_checkListBox.ClearSelected();
            ques15_checkListBox.ClearSelected();
            ques16_comBox.SelectedIndex = 0;
            ques18_comBox.SelectedIndex = 0;
            ques19_comBox.SelectedIndex = 0;
        }

        private void fileOpen()
        {
            try
            {
                mApp = new Excel.Application();
                mApp.Visible = true;
                //myBook = mApp.Workbooks.Add(filepath);
                myBook = mApp.Workbooks.Open(filepath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true, Type.Missing, Type.Missing);
                sheets = myBook.Sheets;

                mySheet = sheets[1];
                if (mySheet == null)
                {
                    Console.WriteLine("没有工作簿");
                    return;
                }
                mySheet.Activate();

                rowCount = mySheet.UsedRange.Rows.Count + 1;
                Console.WriteLine("加载时行数：" + rowCount);
                status_textBox.Text = "文件初始化成功";

                insertTitle();
               
            }catch(Exception ex)
            {
                status_textBox.Text = ex.Message;
                mApp.Quit();
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //fileOpen();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            closeFile();
        }


        ///// <summary>
        ///// 切换选项卡
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    //insertTitle();
        //    if (button1.Text == "下一页")
        //    {
        //        tabControl1.SelectedIndex = 1;
        //    }
        //    else
        //    {
        //        tabControl1.SelectedIndex = 0;
        //    }
        //}

        /// <summary>
        /// 提交信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (filepath == null)
            {
                MessageBox.Show("未输入文件名");
                return;
            }

            try
            {
                insertInfo();
            }catch(Exception ex)
            {
                Console.WriteLine("写入异常：" + ex.Message);
                status_textBox.Text = "写入异常：" + ex.Message;
            }
            //try
            //{

            //    mySheet.Cells[rowCount, 1] = filepath;
            //    mySheet.Cells[rowCount, 2] = "lingchen";
            //    mySheet.Cells[rowCount, 3] = "nihao";
            //    rowCount++;
            //}
            //catch(Exception ex)
            //{
            //    Console.WriteLine("异常：" + ex.Message);
            //    status_textBox.Text = ex.Message;
            //}

        }

        public void insertInfo()
        {

            question.school = Convert.ToInt32(school_textBox.Text);
            question.area = Convert.ToInt32(area_textBox.Text);
            question.grade = Convert.ToInt32(grade_textBox.Text);

            //基本信息
            mySheet.Cells[rowCount, 1] = rowCount-1;
            mySheet.Cells[rowCount, 2] = question.school;
            mySheet.Cells[rowCount, 3] = question.area;
            mySheet.Cells[rowCount, 4] = question.grade;
            mySheet.Cells[rowCount, 5] = question.ques1;
            
            for(int i = 0; i < 10; i++)
            {
                mySheet.Cells[rowCount, 6 + i] = question.ques2_base[i];
            }           
            mySheet.Cells[rowCount, 16] = question.ques2_other;

            if (question.ques3 != 4)
            {
                mySheet.Cells[rowCount, 17] = question.ques3;
            }
            else
            {
                mySheet.Cells[rowCount, 17] = question.ques3_other;
            }

            mySheet.Cells[rowCount, 18] = question.ques4;
            mySheet.Cells[rowCount, 19] = question.ques5;

            mySheet.Cells[rowCount, 20] = question.ques6;
            mySheet.Cells[rowCount, 21] = question.ques7;

            if(question.ques8 != 5)
            {
                mySheet.Cells[rowCount, 22] = question.ques8;
            }
            else
            {
                mySheet.Cells[rowCount, 22] = question.ques8_other;
            }
         

            mySheet.Cells[rowCount, 23] = question.ques9;
            mySheet.Cells[rowCount, 24] = question.ques10;

            mySheet.Cells[rowCount, 25] = question.ques11;
            for(int i = 0; i < 10; i++)
            {
                mySheet.Cells[rowCount, 26+i] = question.ques12_base[i];
            }
            mySheet.Cells[rowCount, 36] = question.ques12_other;


            mySheet.Cells[rowCount, 37] = question.ques13;
            for(int i = 0; i < 8; i++)
            {
                mySheet.Cells[rowCount, 38+i] = question.ques14[i];
            }

            for(int i = 0; i < 5; i++)
            {
                mySheet.Cells[rowCount, 46+i] = question.ques15[i];
            }

            mySheet.Cells[rowCount, 51] = question.ques16;

            for(int i = 0; i < 10; i++)
            {
                mySheet.Cells[rowCount,52+i] = question.ques17[i];
            }


            mySheet.Cells[rowCount, 62] = question.ques18;
            mySheet.Cells[rowCount, 63] = question.ques19;



            //其他信息

            for(int i = 0; i < 96; i++)
            {
                mySheet.Cells[rowCount, 64 + i] = question.ohterQues[i];
            }

            rowCount++; 
           
        }




        private void insertTitle()
        {
            string array = "1,2.1,2.2,2.3,2.4,2.5,2.6,2.7,2.8,2.9,2.10,2.11,3,4,5,6,7,8,9,10,11,12.1,12.2,12.3,12.4,12.5,12.6,12.7,12.8,12.9,12.10,12.11,13,14.1,14.2,14.3,14.4,14.5,14.6,14.7,14.8,15.1,15.2,15.3,15.4,15.5,16,17.1,17.2,17.3,17.4,17.5,17.6,17.7,17.8,17.9,17.10,18,19,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96";

            //string temp = array.Replace(Convert.ToChar(" "), Convert.ToChar(','));

            Console.WriteLine(array);
            string[] titles = array.Split(',');


            //char empty = Convert.ToChar(" ");
            //string[]  titleList = array.Split(Convert.ToChar(" "));
            //for(int i = 0; i < titleList.Length; i++)
            //{
            //    Console.WriteLine(titleList[1]);
            //}
            mySheet.Cells[1, 2] = "学校";
            mySheet.Cells[1, 3] = "区域";
            mySheet.Cells[1, 4] = "年纪";
            for(int i = 0; i < titles.Length; i++)
            {
                mySheet.Cells[1, 5 + i] = titles[i];
            }
            
            //mySheet.Cells[1, 5] = "1";
            //mySheet.Cells[1, 6] = "2.1";
            //mySheet.Cells[1, 7] = "2.2";
            //mySheet.Cells[1, 8] = "2.3";
            //mySheet.Cells[1, 9] = "2.4";


        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

            //switch (tabControl1.SelectedIndex)
            //{
            //    case 0:
            //        button2.Visible = false;
            //        button1.Text = "下一页";
            //        break;
            //    case 1:
            //        button2.Visible = true;
            //        button1.Text = "上一页";
            //        break;
            //    case 2:
            //        button2.Visible = true;
            //        break
                     
            //}
        }


        private void ques3_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(ques3_comBox.Text == other)
            {
                ques3_text_more.Visible = true;
                //question.ques3 = -1;  
            }
            else
            {
                ques3_text_more.Visible = false;
            }
            
                question.ques3 = ques3_comBox.SelectedIndex + 1;
                //ques3_text_more.Visible = false;
            
        }

        private void ques8_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(ques8_comBox.Text == other)
            {
                ques8_text_more.Visible = true;
                //question.ques8 = -1;
            }
            else
            {
                ques8_text_more.Visible = false;
              
            }
            question.ques8 = ques8_comBox.SelectedIndex + 1;
        }

        private void ques12_checkListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            //List<int> ques12_Sel = new List<int>();
            for(int i = 0; i < ques12_checkListBox.Items.Count; i++)
            {
                if (ques12_checkListBox.GetItemChecked(i))
                {
                    //ques12_Sel.Add(i + 1);
                    question.ques12_base[i] = 1;
                }
                else
                {
                    question.ques12_base[i] = 0;
                }
            }
            //question.ques12_base = ques12_Sel;
        }

        private void ques2_more_text_TextChanged(object sender, EventArgs e)
        {
            question.ques2_other = ques2_more_text.Text;
        }

        private void ques3_text_more_TextChanged(object sender, EventArgs e)
        {
            question.ques3_other = ques3_text_more.Text;
        }

        private void ques8_text_more_TextChanged(object sender, EventArgs e)
        {
            question.ques8_other = ques8_text_more.Text;
        }

        private void ques12_text_more_TextChanged(object sender, EventArgs e)
        {
            question.ques12_other = ques12_text_more.Text;
        }

        private void ques1_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques1 = ques1_comBox.SelectedIndex + 1;
        }

        private void ques2_checkListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List<int> ques2_Sel = new List<int>();
            for (int i = 0; i < ques2_checkListBox.Items.Count; i++)
            {
                if (ques2_checkListBox.GetItemChecked(i))
                {
                    //ques2_Sel.Add(i + 1);
                    question.ques2_base[i] = 1;
                }
                else
                {
                    question.ques2_base[i] = 0;
                }
            }
            //question.ques2_base = ques2_Sel;
        }

        private void ques4_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques4 = ques4_comBox.SelectedIndex + 1;
        }

        private void ques5_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques5 = ques5_comBox.SelectedIndex + 1;
        }

        private void ques6_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques6 = ques6_comBox.SelectedIndex + 1;
        }

        private void ques7_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques7 = ques7_comBox.SelectedIndex + 1;
        }

        private void ques9_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques9 = ques9_comBox.SelectedIndex + 1;
        }

        private void ques10_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques10 = ques10_comBox.SelectedIndex + 1;
        }

        private void ques11_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques11 = ques11_comBox.SelectedIndex + 1;
        }

        private void ques13_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques13 = ques13_comBox.SelectedIndex + 1;
        }

        private void ques14_checkListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List<int> ques14_Sel = new List<int>();
            for (int i = 0; i < ques14_checkListBox.Items.Count; i++)
            {
                if (ques14_checkListBox.GetItemChecked(i))
                {
                    //ques14_Sel.Add(i + 1);
                    question.ques14[i] = 1;
                }
                else
                {
                    question.ques14[i] = 0;
                }
            }
            //question.ques14 = ques14_Sel;
        }

        private void ques15_checkListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List<int> ques15_Sel = new List<int>();
            for (int i = 0; i < ques15_checkListBox.Items.Count; i++)
            {
                if (ques15_checkListBox.GetItemChecked(i))
                {
                    //ques15_Sel.Add(i + 1);
                    question.ques15[i] = 1;
                }
                else
                {
                    question.ques15[i] = 0;
                }
            
            }
            //question.ques15 = ques15_Sel;
        }

        private void ques16_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques16 = ques16_comBox.SelectedIndex + 1;
        }

        private void ques17_checkListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //List<int> ques17_Sel = new List<int>();
            for (int i = 0; i < ques17_checkListBox.Items.Count; i++)
            {
                if (ques17_checkListBox.GetItemChecked(i))
                {
                    //ques17_Sel.Add(i + 1);
                    question.ques17[i] = 1;
                }
                else
                {
                    question.ques17[i] = 0;
                }
            }
            //question.ques17 = ques17_Sel;
        }

        private void ques18_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques18 = ques18_comBox.SelectedIndex + 1;
        }

        private void ques19_comBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            question.ques19 = ques19_comBox.SelectedIndex + 1;
        }

        private void school_textBox_TextChanged(object sender, EventArgs e)
        {
            question.school = Convert.ToInt32(school_textBox.Text);
        }

        private void area_textBox_TextChanged(object sender, EventArgs e)
        {
            question.area = Convert.ToInt32(area_textBox.Text);
        }

        private void grade_textBox_TextChanged(object sender, EventArgs e)
        {
            question.grade = Convert.ToInt32(grade_textBox.Text);
        }

        private void file_path_textBox_TextChanged(object sender, EventArgs e)
        {
            filepath = file_path_textBox.Text;
           
            Console.WriteLine("文件名变更:"+filepath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fileOpen();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            closeFile();
        }

        public void closeFile()
        {
            try
            {
                if (myBook != null)
                {
                    mApp.DisplayAlerts = false;
                    Console.WriteLine("关闭：" + filepath);
                    myBook.SaveAs(filepath, Missing.Value, Missing.Value, Missing.Value, false, Missing.Value, XlSaveAsAccessMode.xlExclusive, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    mApp.DisplayAlerts = true;
                    myBook.Close(true, filepath, Missing.Value);
                    myBook = null;
                    mApp.Quit();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("异常：" + ex.Message);
            }
        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = (ComboBox)sender;
            string name = combox.Name;
            int  nameIndex  =Convert.ToInt32(name.Substring(8));
            question.ohterQues[nameIndex - 1] = combox.SelectedIndex + 1;
        }
    }
}
