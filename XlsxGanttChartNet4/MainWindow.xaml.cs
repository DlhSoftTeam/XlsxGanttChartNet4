using DlhSoft.Windows.Controls;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace XlsxGanttChartNet4
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            for (var i = 1; i <= 10; i++)
                GanttChartDataGrid.Items.Add(new GanttChartItem { Content = "Task " + i, Start = DateTime.Today.AddDays(i), Finish = DateTime.Today.AddDays(i * i) });
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog { Filter = "Excel files|*.xlsx", DefaultExt = ".xlsx", Title = "Save Excel file" };
            if (dialog.ShowDialog() == true)
            {
                var projectXml = GanttChartDataGrid.GetProjectXml();
                var excelBytes = ProjectManagementXlsx.Adapter.GetExcelBytes(projectXml);
                using (var stream = dialog.OpenFile())
                {
                    stream.Write(excelBytes, 0, excelBytes.Length);
                }
            }
        }

        private void LoadButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filter = "Excel files|*.xlsx", Title = "Open Excel file" };
            if (dialog.ShowDialog() == true)
            {
                using (var stream = dialog.OpenFile())
                {
                    var excelBytes = new byte[stream.Length];
                    stream.Read(excelBytes, 0, excelBytes.Length);
                    var projectXml = ProjectManagementXlsx.Adapter.GetProjectXml(excelBytes);
                    GanttChartDataGrid.LoadProjectXml(projectXml);
                }
            }
        }
    }
}
