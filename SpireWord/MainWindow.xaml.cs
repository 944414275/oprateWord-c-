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
using Spire.Doc;
using Spire.Doc.Documents;
using SpireWord.modul;
using Spire.Doc.Fields;
using Microsoft.Win32;


namespace SpireWord
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        int columeCount=4;
        int dataRowCount;
        Document wordDocument = new Document();
        Student exportStu = new Student();
        Spire.Doc.Documents.Paragraph para =null;
        Spire.Doc.Fields.TextRange TR =null;
        List<Student> listStu = new List<Student>();
        string[] strHeader = { "Leader", "Name", "Score", "StuClass"};

        public MainWindow()
        {
            InitializeComponent();
        }

        
        private void ExprtWord(object sender, RoutedEventArgs e)
        {
            //initial data source
            listStu = exportStu.getStuData();
            dataRowCount = listStu.Count+1;
            string path;

            //creat document、settion and table
            Spire.Doc.Section stuInfoSection =wordDocument.AddSection();
            Spire.Doc.Table stuTable =stuInfoSection.AddTable(true);
            stuTable.ResetCells(dataRowCount, columeCount);

            //set header 
            Spire.Doc.TableRow row =stuTable.Rows[0];//initial TableRow
            row.IsHeader = true;
            for(int i=0;i<columeCount;i++)
            {
                para=row.Cells[i].AddParagraph();
                TR = para.AppendText(strHeader[i]);
            }

            //fill data
            for (int i = 1; i < dataRowCount; i++)
            {
                
                for (int j = 0; j < columeCount; j++)
                {
                    para = stuTable.Rows[i].Cells[j].AddParagraph();
                    TR = para.AppendText(listStu[i-1].GetValue(listStu[i-1].propertyIndex[j]));
                }
            }

            //
            
            //save file
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Documents(*.docx)|*.docx";
            var res = saveFileDialog.ShowDialog();
            if (res != true) return;
            path = saveFileDialog.FileName;
            MessageBox.Show(path);
            
            wordDocument.SaveToFile(@path, FileFormat.Docx);
            wordDocument.Close();
        }
    }
}
