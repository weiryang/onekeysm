using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OneKeySendMail
{
    class Program
    {
        static void Main(string[] args) 
        {
            if (args.Length != 5)
            {
                Console.WriteLine("");
                Console.WriteLine("参数不足,完整格式如下：");
                Console.WriteLine("OneKeySendMail.exe 杨威 F:\\\\xlh_WorkAbout\\\\工作日报-杨威.xls yangwei@hljxlh.net xxxxxx 0");
                Console.WriteLine("第一个参数:发送人姓名");
                Console.WriteLine("第二个参数:日报文件的全路径");
                Console.WriteLine("第三个参数:发送用邮箱");
                Console.WriteLine("第四个参数:发送用邮箱的密码 ");
                Console.WriteLine("第五个参数:测试用参数，0 正常发送到三个邮箱，1 只发送到第三个参数的邮箱中 ");
                Console.WriteLine("");
                Console.WriteLine("自动根据当前日期和第一个参数生成主题，如:基地-学术部-杨威-20160901 ");
                Console.WriteLine("正文为 “RT”，如题 ");
                Console.WriteLine("最后一个参数为“0”时，将发送到\"基地考核邮箱\",\"戚爱滨\",\"孙文弼\"三个邮箱 ");
                Console.WriteLine("PS:程序现暂不能实现定时发送，并且执行的时候必须在17：30以后，可以在windows系统里定义任务计划，如在17：35执行命令即可 ");
                Console.WriteLine("正常发送的日报邮件，可以在发件箱中找到");
                Console.WriteLine("本程序目的只为了方便发送邮件，日报内容仍需要及时手动更新");
                return;
            }

            //主题定义,生成类似“基地-移动互联软件研发部-杨威-20160829”格式的字符串
            string sendName = args[0];
            string sub = "基地-学术部-" + sendName + "-";
            sub += DateTime.Now.ToString("yyyyMMdd");
            //string filepath = "F:\\xlh_WorkAbout\\工作日报-杨威.xls";
            string filepath = args[1];
            string ccAddress = args[2];
            string ccPwd = args[3];
            string ton = args[4];//test or no , only 0,or 1
            if (ton != "0" && ton !="1")
            {
                Console.WriteLine("最后一个参数只能为 0 或 1");
                return; 
            }


            //Console.WriteLine(sub);
            //return;
                 
            //Mail mail = new Mail("a_1234_a@163.com", "杨威", "tyblxl@163.com", "Mr.Yang", sub, "RT", "smtp.163.com", 25, "a_1234_a", "1122334455", true, filepath);
            Mail mail = new Mail(ccAddress, sendName, ccAddress, sendName, sub, "RT", 
                "smtp.ym.163.com", 25, ccAddress, ccPwd, true, filepath);
            bool sr = mail.Send(ccAddress,ton);
            if (sr)
            {
                Console.WriteLine("----------------------------------------");
                Console.WriteLine("华育兴业--工作邮件一键发送");
                Console.WriteLine("----------------------------------------");
                Console.WriteLine("邮件成功发送！");
            }
            else
            {
                Console.WriteLine("邮件发送失败！");
            }

            //Console.WriteLine("请按任意键继续");
            //Console.ReadKey();
        }
    }
}
