using System;
using System.Collections.Generic;
using System.Configuration;
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
            log4net.Config.XmlConfigurator.Configure();
            ViData.DMHelper.Instance.ExportMapping();            
        }
        public static string filepath = ConfigurationManager.AppSettings["basepath"];

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        void ExportFiles()
        {
            for(int i = 1;i<=2; i++)
            {
                string path = filepath + i;
                //数据统计_2017年4月.xlsx
                string filename = string.Format("数据统计_{0}年{1}月.xlsx", DateTime.Now.Year, DateTime.Now.Month);

            }
        }
    }
}
