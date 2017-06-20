using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PSQTransform
{
    class Questionnaire
    {

        public int school = 1;
        public int area = 1;
        public int grade = 1;


        public int ques1 = 1;//默认1:男
        public int[] ques2_base = new int[10];//爱好
        public string ques2_other ="0";
        public int ques3 = 1;//居住区
        public string ques3_other = "0";
        public int ques4 = 1;//玩耍场所
        public int ques5 = 1;//图书馆
        public int ques6 = 1;
        public int ques7 = 1;
        public int ques8 = 1;
        public string ques8_other = "0";
        public int ques9 = 1;
        public int ques10 = 1;
        public int ques11 = 1;
        public int[] ques12_base = new int[10];
        public string ques12_other = "0";
        public int ques13 = 1;
        public int[] ques14 = new int[8];
        public int[] ques15 = new int[5];
        public int ques16 = 1;
        public int[] ques17 = new int[10];
        public int ques18 = 1;
        public int ques19 = 1;

        public int[] ohterQues = new int[96];

        public void toString()
        {
            Console.WriteLine(ques1 + "," + ques2_other+","+ques3_other+","+ques4);
        }
    }
}
