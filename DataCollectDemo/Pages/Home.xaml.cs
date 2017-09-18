using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DataCollectDemo.ViewModels;
using Application = System.Windows.Application;
using Cursor = System.Windows.Input.Cursor;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using UserControl = System.Windows.Controls.UserControl;

namespace DataCollectDemo.Pages
{
    /// <summary>
    /// Interaction logic for Home.xaml
    /// </summary>
    public partial class Home : UserControl
    {
        public Home()
        {
            ViewModel = new HomeViewModel();
            this.DataContext = ViewModel;
            InitializeComponent();
        }

        public HomeViewModel ViewModel { get; set; }

        private void Output_Click(object sender, RoutedEventArgs e)
        {
            string savedPath = SaveDialogResult();
            if (savedPath == null) return;

            ViewModel.SavedPath = savedPath;

            ViewModel.SelectedPath = FilePath.Text;
            ViewModel.BookColumn = CharToInt(BookColumn.Text);
            ViewModel.ClassColumn = CharToInt(ClassColumn.Text);
            ViewModel.DepartColumn = CharToInt(DepartColumn.Text);
            //ViewModel.SubjectColumn = CharToInt(SubjectColumn.Text);

            var act = new Action<string>(NotifyUser);
            var tsk = new Task(ViewModel.StartCollect, act);
            tsk.Start();
            tsk.ContinueWith(task => NotifyUser(DateTime.Now + "数据提取完成！！！"));
        }

        private int CharToInt(string text)
        {
            return text.ToUpper()[0] - 'A' + 1;
        }

        private void SelectedPath_Click(object sender, RoutedEventArgs e)
        { 
            FilePath.Text = GetDialogResult();
        }

        private string SaveDialogResult()
        {
            using (SaveFileDialog openFileDialog = new SaveFileDialog
            {
                Filter = "Excel File|*.xlsx;*.xls",
                Title = "请选择excel文件",
                InitialDirectory = @"Documents"
            })
            {
                try
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                        return openFileDialog.FileName;
                    return null;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    NotifyUser(e.Message);
                    Application.Current.Shutdown();
                }
            }
            return String.Empty;
        }

        private string GetDialogResult()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel File|*.xls;*.xlsx",
                Title = "请选择excel文件",
                InitialDirectory = @"Documents"
            };
            try
            {
                bool? result = openFileDialog.ShowDialog();

                if (result == true)
                {
                    return openFileDialog.FileName;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                NotifyUser(exception.Message);
                Application.Current.Shutdown();
            }

            return String.Empty;
        }

        private void NotifyUser(string message)
        {
            //NotifyZone.Text = message;
            ViewModel.Message = message;
        }
    }
}
