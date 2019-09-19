using System;
using System.Collections.Generic;
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

namespace OperateWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 数据实体类
        /// </summary>
        public class Student
        {
            public string Name;//姓名
            public int Score;//成绩
            public string StuClass;//班级
            public string Leader;//班主任
        }
        /// <summary>
        /// 动态创建table到word
        /// </summary>
        protected void CreateTableToExcel()
        {
            Word.Application app = null;
            Word.Document doc = null;
            try
            {
                //构造数据
                List<Student> datas = new List<Student>();
                datas.Add(new Student { Leader = "小李", Name = "张三", Score = 498, StuClass = "一班" });
                datas.Add(new Student { Leader = "陈飞", Name = "李四", Score = 354, StuClass = "二班" });
                datas.Add(new Student { Leader = "陈飞", Name = "小红", Score = 502, StuClass = "二班" });
                datas.Add(new Student { Leader = "王林", Name = "丁爽", Score = 566, StuClass = "三班" });
                var cate = datas.GroupBy(s => s.StuClass);

                int rows = cate.Count() + 1;//表格行数加1是为了标题栏
                int cols = 5;//表格列数
                object oMissing = System.Reflection.Missing.Value;
                app = new Word.Application();//创建word应用程序
                doc = app.Documents.Add();//添加一个word文档

                //输出大标题加粗加大字号水平居中
                app.Selection.Font.Bold = 700;
                app.Selection.Font.Size = 16;
                app.Selection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                app.Selection.Text = "班级成绩统计单";

                //换行添加表格
                object line = Word.WdUnits.wdLine;
                app.Selection.MoveDown(ref line, oMissing, oMissing);
                app.Selection.TypeParagraph();//换行
                Word.Range range = app.Selection.Range;
                Word.Table table = app.Selection.Tables.Add(range, rows, cols, ref oMissing, ref oMissing);

                //设置表格的字体大小粗细
                table.Range.Font.Size = 10;
                table.Range.Font.Bold = 0;

                //设置表格标题
                int rowIndex = 1;
                table.Cell(rowIndex, 1).Range.Text = "班级";
                table.Cell(rowIndex, 2).Range.Text = "姓名";
                table.Cell(rowIndex, 3).Range.Text = "成绩";
                table.Cell(rowIndex, 4).Range.Text = "人数";
                table.Cell(rowIndex, 5).Range.Text = "班主任";

                //循环数据创建数据行
                rowIndex++;
                foreach (var i in cate)
                {
                    table.Cell(rowIndex, 1).Range.Text = i.Key;//班级
                    table.Cell(rowIndex, 4).Range.Text = i.Count().ToString();//人数
                    table.Cell(rowIndex, 5).Range.Text = i.First().Leader;//班主任
                    table.Cell(rowIndex, 2).Split(i.Count(), 1);//分割名字单元格
                    table.Cell(rowIndex, 3).Split(i.Count(), 1);//分割成绩单元格

                    //对表格中的班级、姓名，成绩单元格设置上下居中
                    table.Cell(rowIndex, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(rowIndex, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(rowIndex, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    //构建姓名，成绩数据
                    foreach (var x in i)
                    {
                        table.Cell(rowIndex, 2).Range.Text = x.Name;
                        table.Cell(rowIndex, 3).Range.Text = x.Score.ToString();
                        rowIndex++;
                    }
                }

                //导出到文件
                string newFile = DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc";
                string physicNewFile = Server.MapPath(newFile);
                doc.SaveAs(physicNewFile,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();//关闭文档
                }
                if (app != null)
                {
                    app.Quit();//退出应用程序
                }
            }
        }

        protected void CreateTableToExcel2()
        {
            Word.Application app = null;
            Word.Document doc = null;
            try
            {
                //构造数据
                List<Student> datas = new List<Student>();
                datas.Add(new Student { Leader = "小李", Name = "张三", Score = 498, StuClass = "一班" });
                datas.Add(new Student { Leader = "陈飞", Name = "李四", Score = 354, StuClass = "二班" });
                datas.Add(new Student { Leader = "陈飞", Name = "小红", Score = 502, StuClass = "二班" });
                datas.Add(new Student { Leader = "王林", Name = "丁爽", Score = 566, StuClass = "三班" });
                var cate = datas.GroupBy(s => s.StuClass);

                int rows = datas.Count + 1;
                int cols = 5;//表格列数
                object oMissing = System.Reflection.Missing.Value;
                app = new Word.Application();//创建word应用程序
                doc = app.Documents.Add();//添加一个word文档

                //输出大标题加粗加大字号水平居中
                app.Selection.Font.Bold = 700;
                app.Selection.Font.Size = 16;
                app.Selection.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                app.Selection.Text = "班级成绩统计单";

                //换行添加表格
                object line = Word.WdUnits.wdLine;
                app.Selection.MoveDown(ref line, oMissing, oMissing);
                app.Selection.TypeParagraph();//换行
                Word.Range range = app.Selection.Range;
                Word.Table table = app.Selection.Tables.Add(range, rows, cols, ref oMissing, ref oMissing);

                //设置表格的字体大小粗细
                table.Range.Font.Size = 10;
                table.Range.Font.Bold = 0;

                //设置表格标题
                int rowIndex = 1;
                table.Cell(rowIndex, 1).Range.Text = "班级";
                table.Cell(rowIndex, 2).Range.Text = "姓名";
                table.Cell(rowIndex, 3).Range.Text = "成绩";
                table.Cell(rowIndex, 4).Range.Text = "人数";
                table.Cell(rowIndex, 5).Range.Text = "班主任";

                //循环数据创建数据行
                rowIndex++;
                foreach (var i in cate)
                {
                    int moveCount = i.Count() - 1;//纵向合并行数
                    if (moveCount.ToString() != "0")
                    {
                        table.Cell(rowIndex, 1).Merge(table.Cell(rowIndex + moveCount, 1));//合并班级
                        table.Cell(rowIndex, 4).Merge(table.Cell(rowIndex + moveCount, 4));//合并人数
                        table.Cell(rowIndex, 5).Merge(table.Cell(rowIndex + moveCount, 5));//合并班主任
                    }
                    //写入合并的数据并垂直居中
                    table.Cell(rowIndex, 1).Range.Text = i.Key;
                    table.Cell(rowIndex, 4).Range.Text = i.Count().ToString();
                    table.Cell(rowIndex, 5).Range.Text = i.First().Leader;
                    table.Cell(rowIndex, 1).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(rowIndex, 4).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    table.Cell(rowIndex, 5).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    //构建姓名，成绩数据
                    foreach (var x in i)
                    {
                        table.Cell(rowIndex, 2).Range.Text = x.Name;
                        table.Cell(rowIndex, 3).Range.Text = x.Score.ToString();
                        rowIndex++;
                    }
                }
                //导出到文件
                string newFile = DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc";
                string physicNewFile = Server.MapPath(newFile);
                doc.SaveAs(physicNewFile,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();//关闭文档
                }
                if (app != null)
                {
                    app.Quit();//退出应用程序
                }
            }
        }
    }


    /// <summary>
    /// 数据实体类
    /// </summary>
    public class Student
    {
        public string Name;//姓名
        public int Score;//成绩
        public string StuClass;//班级
        public string Leader;//班主任
    }
}
