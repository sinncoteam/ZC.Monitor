using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ZC.Utils;

namespace ZC.Client
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();                     
        }
        public static string filepath = ConfigurationManager.AppSettings["basepath"];

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            initData();
        }

        void initData()
        {
            log4net.Config.XmlConfigurator.Configure();
            ViData.DMHelper.Instance.ExportMapping();
            setColname();
        }

        void ExportFiles()
        {
            for(int i = 1;i<=2; i++)
            {
                string path = filepath + i;
                if (Directory.Exists(path))
                {
                    //数据统计_2017年4月.xlsx
                    string filename = string.Format("\\数据统计_{0}年{1}月.xlsx", DateTime.Now.Year, DateTime.Now.Month);
                    path += filename;
                    showText("开始读取文件" + path);
                    ExcelHelper exc = new ExcelHelper(path);
                    var dt = exc.ExcelToDataTable("sheet1", false,5);
                    int n = 0;
                    foreach(DataColumn col in dt.Columns)
                    {
                        if (colList.Count > n)
                        {
                            var dcitem = colList[n];
                            col.ColumnName = dcitem;
                        }
                        else
                        {
                            break;
                        }
                        n++;
                    }
                    dg_result.AutoGenerateColumns = true;
                    dg_result.ItemsSource = dt.DefaultView;
                }
            }
        }
        void showText(string text)
        {
            this.rtb_show.AppendText(text + "\r");
            this.rtb_show.ScrollToEnd();
        }
        private void btn_start_Click(object sender, RoutedEventArgs e)
        {
            ExportFiles();
        }

        public static List<string> colList = new List<string>();
        public static void  setColname()
        {            
            colList.Add("序号");
            colList.Add("检测日期");
            colList.Add("班别");
            colList.Add("班次机组号");
            colList.Add("品名规格");
            colList.Add("检测机号");
            colList.Add("总进瓶数");
            colList.Add("碎瓶剔除瓶数");
            colList.Add("碎瓶剔除率");
            colList.Add("检验总数");
            colList.Add("合格数");
            colList.Add("合格率");
            colList.Add("总不良数");
            colList.Add("总不良率");
            colList.Add("规格尺寸不良总数");
            colList.Add("规格尺寸不良率");
            colList.Add("外观不良总数");
            colList.Add("外观不良率");
            colList.Add("瓶身外径缺陷不良个数");
            colList.Add("占不良比例");
            colList.Add("占检验数比例");
        }
    }
}
