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
using Visifire.Charts;
using ZC.Client.Base;
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
            List<ExcelObject> objList = new List<ExcelObject>();
            for (int i = 1;i<=3; i++)
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
                    foreach(DataRow row in dt.Rows)
                    {
                       
                        string machineId = row[(int)ExcelCols.检测机号].ToString();
                        var item = objList.Where(a => a.MachineId == machineId).FirstOrDefault();
                        if (item != null)
                        {
                            item.TotalCheck += Convert.ToInt32(row[(int)ExcelCols.检验总数]);
                            item.TotalGood += Convert.ToInt32(row[(int)ExcelCols.总合格数]);
                            item.TotalBad += Convert.ToInt32(row[(int)ExcelCols.总不良数]);
                            double good = Math.Round( (double)item.TotalGood * 100 / item.TotalCheck, 2);
                            double bad = Math.Round((double)item.TotalBad * 100 / item.TotalCheck, 2);
                            item.TotalGoodPercent = good;
                            item.TotalBadPercent = bad;
                        }
                        else
                        {
                            ExcelObject obj = new ExcelObject();
                            obj.MachineId = row[(int)ExcelCols.检测机号].ToString();
                            obj.TotalCheck = Convert.ToInt32(row[(int)ExcelCols.检验总数]);
                            obj.TotalGood = Convert.ToInt32(row[(int)ExcelCols.总合格数]);
                            obj.TotalGoodPercent = StringHelper.PercentToInt(row[(int)ExcelCols.总合格率].ToString());
                            obj.TotalBad = Convert.ToInt32(row[(int)ExcelCols.总不良数]);
                            obj.TotalBadPercent = StringHelper.PercentToInt(row[(int)ExcelCols.总不良率].ToString());
                            objList.Add(obj);
                        }
                    }
                    //int n = 0;
                    //foreach(DataColumn col in dt.Columns)
                    //{
                    //    if (colList.Count > n)
                    //    {
                    //        var dcitem = colList[n];
                    //        col.ColumnName = dcitem;
                    //    }
                    //    else
                    //    {
                    //        break;
                    //    }
                    //    n++;
                    //}
                    //dg_result.AutoGenerateColumns = true;
                    //dg_result.ItemsSource = dt.DefaultView;
                }
            }
            int total = 0;
            int totalgood = 0;
            List<string> machineList = new List<string>();
            List<double> resultList = new List<double>();
            foreach(var item in objList)
            {
                machineList.Add(item.MachineId + "号机合格率");
                resultList.Add(item.TotalGoodPercent);
                machineList.Add(item.MachineId + "号机不合格率");
                resultList.Add(item.TotalBadPercent);
                total += item.TotalCheck;
                totalgood += item.TotalGood;                
            }
            
            CreateChartColumn("总体检测情况图", machineList, resultList, "%");

            double totalgoodper = Math.Round((double)totalgood * 100 / total, 2);
            tb_total_yes.Text = totalgood.ToString();
            tb_total_yesper.Text = totalgoodper.ToString() + "%";
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

        public void CreateChartColumn(string name, List<string> valuex, List<double> valuey, string suffix)
        {
            //创建一个图标
            Chart chart = new Chart();

            //设置图标的宽度和高度
            chart.Width = 600;
            chart.Height = 380;
            //chart.Margin = new Thickness(100, 5, 10, 5);
            //是否启用打印和保持图片
            chart.ToolBarEnabled = false;

            //设置图标的属性
            chart.ScrollingEnabled = false;//是否启用或禁用滚动
            chart.View3D = true;//3D效果显示

            //创建一个标题的对象
            Title title = new Title();

            //设置标题的名称
            title.Text = Name;
            title.Padding = new Thickness(0, 10, 5, 0);

            //向图标添加标题
            chart.Titles.Add(title);

            Axis yAxis = new Axis();
            //设置图标中Y轴的最小值永远为0           
            yAxis.AxisMinimum = 0;
            //设置图表中Y轴的后缀          
            yAxis.Suffix = suffix;
            chart.AxesY.Add(yAxis);

            // 创建一个新的数据线。               
            DataSeries dataSeries = new DataSeries();

            // 设置数据线的格式
            dataSeries.RenderAs = RenderAs.StackedColumn;//柱状Stacked


            // 设置数据点              
            DataPoint dataPoint;
            for (int i = 0; i < valuex.Count; i++)
            {
                // 创建一个数据点的实例。                   
                dataPoint = new DataPoint(); 

                // 设置X轴点                    
                dataPoint.AxisXLabel = valuex[i]; 
                //设置Y轴点                   
                dataPoint.YValue = valuey[i];
               
                //添加一个点击事件        
                //dataPoint.MouseLeftButtonDown += new MouseButtonEventHandler(dataPoint_MouseLeftButtonDown);
                //添加数据点                   
                dataSeries.DataPoints.Add(dataPoint);
            }

            // 添加数据线到数据序列。                
            chart.Series.Add(dataSeries);

            //将生产的图表增加到Grid，然后通过Grid添加到上层Grid.           
            
            g_chart.Children.Add(chart);
        }
    }
}
