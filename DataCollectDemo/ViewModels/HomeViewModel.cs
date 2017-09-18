using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DataCollectDemo.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataCollectDemo.ViewModels
{
    public class HomeViewModel : BindableBase
    {
        public HomeViewModel()
        {
            Classes = new Dictionary<string, IList<Book>>();
            //Subjects = new Dictionary<string, IList<string>>();
            Departments = new Dictionary<string, IList<string>>();

            BooksNames = new List<string>();
            //Subjs = new List<string>();
            Deps = new List<string>();

            DefaultParams();
        }

        private void DefaultParams()
        {
            Message = "";
            Semester = "2017-2018-01学期";
            StartedRow = 2;
            Grade = "15级";
        }

        public String SelectedPath { get; set; }

        public int BookColumn { get; set; }

        public int DepartColumn { get; set; }

        //public int SubjectColumn { get; set; }

        public int ClassColumn { get; set; }

        public int StartedRow { get; set; }

        public string SavedPath { get; set; }

        private string _message;
        public String Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        private string _semester;

        public string Semester
        {
            get => _semester;
            set => SetProperty(ref _semester, value);
        }
        private string _grade;

        public string Grade
        {
            get => _grade;
            set => SetProperty(ref _grade, value);
        }

        //专业-班级字典
        //public Dictionary<string, IList<string>> Subjects { get; set; }

        //学院-班级字典
        //删改后取消专业层
        public Dictionary<string, IList<string>> Departments { get; set; }
        //班级-图书字典
        public Dictionary<string, IList<Book>> Classes { get; set; }
        //专业列表
        //public IList<string> Subjs { get; set; }
        //学院列表
        public IList<string> Deps { get; set; }

        public IList<string> BooksNames { get; set; }

        public void StartCollect(object obj)
        {
            Action<string> notify = (Action<string>)obj;

            notify("开始读取原始文件。。。。");
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook xlWorkbook = excelApp.Workbooks.Open(SelectedPath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int row = StartedRow;
            notify("正在读取原始文件。。。。");
            while (true)
            {
                //读取专业
                //string classes = GetCellsValue(Subject.Subject, xlRange, row);
                //读取班级
                string classes = GetCellsValue(Subject.Classes, xlRange, row);
                //读取书名
                string booksName = GetCellsValue(Subject.Book, xlRange, row);
                //读取学院
                string depart = GetCellsValue(Subject.Department, xlRange, row);

                //学院为null则认为已读到最后一行，返回
                if (depart == null) break;

                //填充学院列表及字典
                //if (!Departments.ContainsKey(depart))
                //{
                //    Departments.Add(depart, new List<string>());
                //    Deps.Add(depart);
                //}
                //if (!Departments[depart].Contains(classes))
                //{
                //    Departments[depart].Add(classes);
                //}

                //填充学院列表及字典
                if (!Departments.ContainsKey(depart))
                {
                    Departments.Add(depart, new List<string>());
                    Deps.Add(depart);
                }
                if (!Departments[depart].Contains(classes))
                {
                    Departments[depart].Add(classes);
                }

                if (!BooksNames.Contains(booksName))
                {
                    BooksNames.Add(booksName);
                }

                if (!Classes.ContainsKey(classes))
                {
                    Classes[classes] = new List<Book> { new Book { Name = booksName } };
                }
                else
                {
                    bool exist = false;
                    foreach (var abook in Classes[classes])
                    {
                        if (abook.Name == booksName)
                        {
                            abook.Ordered++;
                            exist = true;
                            break;
                        }
                    }
                    if (!exist)
                        Classes[classes].Add(new Book { Name = booksName });
                }
                row++;
            }

            //释放资源
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);


            //输出
            notify("读取完成，正在写入......");
            ExportToExcel();

            Debug.Print("完成！！！！");
            Console.WriteLine("完成！！");
        }

        private void ExportToExcel()
        {
            Excel._Application excelApp = new Excel.Application();

            try
            {

                Excel._Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                //遍历学院
                foreach (var dep in Deps)
                {
                    Excel._Worksheet worksheet = workbook.Worksheets.Add(Type.Missing);

                    worksheet.Name = dep;
                    int relativeRow = 1;

                    // 遍历专业
                    //foreach (var subj in Departments[dep])
                    //{
                    //遍历学院下的班级
                    foreach (var classes in Departments[dep])
                    {
                        //excelApp.Columns.ColumnWidth = 15;
                        PrintHeader(worksheet, relativeRow, dep, classes);
                        PrintSummary(worksheet, relativeRow + 1);
                        PrintContent(worksheet, relativeRow + 2, Classes[classes]);

                        relativeRow = relativeRow + 3 + Classes[classes].Count;
                    }
                    //}
                    //Marshal.ReleaseComObject(worksheet);
                }

                workbook.SaveAs(SavedPath);
                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        private void PrintHeader(Excel._Worksheet worksheet, int row, string department, string classes)
        {
            worksheet.Cells[row, 1].Value2 = "部门";
            worksheet.Cells[row, 2].Value2 = department;
            worksheet.Cells[row, 3].Value2 = classes;
            worksheet.Cells[row, 4].Value2 = Semester;
            worksheet.Cells[row, 10].Value2 = Grade;

            Excel.Range range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 10]];
            range.Font.Bold = true;
            CommonRangeSettings(range);
            range = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row, 4]];
            range.ColumnWidth = 19.8;
        }

        private void PrintSummary(Excel._Worksheet worksheet, int row)
        {
            worksheet.Cells[row, 1].Value2 = "序号";
            worksheet.Cells[row, 2].Value2 = "教材ISBN";
            worksheet.Cells[row, 3].Value2 = "教材名称";
            worksheet.Cells[row, 4].Value2 = "出版社";
            worksheet.Cells[row, 5].Value2 = "作者";
            worksheet.Cells[row, 6].Value2 = "定价";
            worksheet.Cells[row, 7].Value2 = "订数";
            worksheet.Cells[row, 8].Value2 = "领数";
            worksheet.Cells[row, 9].Value2 = "签名";
            worksheet.Cells[row, 10].Value2 = "备注";

            Excel.Range range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 10]];
            range.RowHeight = 30;
            range.Font.Color = 10;
            CommonRangeSettings(range);
        }

        private void PrintContent(Excel._Worksheet worksheet, int row, IList<Book> books)
        {
            for (int i = 0; i < books.Count; i++)
            {
                worksheet.Cells[row, 1].Value2 = i + 1;
                worksheet.Cells[row, 3].Value2 = books[i].Name;//.Split('-')[0];
                worksheet.Cells[row, 7].Value2 = books[i].Ordered;

                Excel.Range range = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 10]];
                range.RowHeight = 25;
                range.WrapText = true;
                CommonRangeSettings(range);
                row++;
            }

        }

        private void CommonRangeSettings(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.Borders.LineStyle = 1;
        }

        private string GetCellsValue(Subject sub, Excel.Range range, int row)
        {


            return ReadCell<string>(range, row, GetRow(sub));
        }

        private int GetRow(Subject sub)
        {
            switch (sub)
            {
                case Subject.Book: return BookColumn;
                case Subject.Department: return DepartColumn;
                case Subject.Classes: return ClassColumn;
                //case Subject.Subject: return SubjectColumn;
                default: return 0;
            }
        }

        private T ReadCell<T>(Excel.Range range, int row, int col)
        {
            try
            {
                return (T)range.Cells[row, col].Value2.ToString();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return default(T);
        }
    }

    public enum Subject
    {
        Book, Department, Classes //, Subject
    }
}
