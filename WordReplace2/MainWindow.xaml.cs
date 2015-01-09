using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.IO;
using System.Data;
using Excel;
using Novacode;

namespace WordReplace
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<UserDoc> inputDocFiles = new ObservableCollection<UserDoc>();
        private UserExcel inputExcelFile;
        private DataSet ds;
        private string outPath;
        private FileHandler fh = new FileHandler();
        private Boolean IsOutputPathSelected = false;
        
        public MainWindow()
        {

            InitializeComponent();
            FileListBox.ItemsSource = inputDocFiles;


        }

        private void DropBox_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[]) e.Data.GetData(DataFormats.FileDrop);
                HandleAddFiles(files);
            }
        }

        private void HandleAddFiles(string[] files)
        {
            foreach (string f in files)
            {
                try
                {
                    UserInputFile newFile = fh.Handle(f);
                    UserDoc newDoс = newFile as UserDoc;

                    if (newDoс != null)
                    {
                        inputDocFiles.Add(newDoс);
                    }
                    else
                    {
                        UserExcel newExcel = newFile as UserExcel;
                        if (newExcel != null)
                        {
                            inputExcelFile = newExcel;
                        }
                    }
                    CheckIsRunEnabled();
                    
                }
                catch (ArgumentException)
                {

                }
                
            }
        }

        private void CheckIsRunEnabled()
        {
            if (inputExcelFile != null && inputDocFiles.Count > 0 && IsOutputPathSelected)
            {
                RunButton.IsEnabled = true;
            }
        }

        private void DropBox_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Filter = "Office documents (.xlsx,.docx)|*.xlsx;*.docx";
            dlg.Multiselect = true;
            var result = dlg.ShowDialog();
            if (result == true)
            {
                string[] files = dlg.FileNames;
                HandleAddFiles(files);
            }
        }

        private void Label_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var folderDialog = new Gat.Controls.OpenDialogView();
            var vm = (Gat.Controls.OpenDialogViewModel)folderDialog.DataContext;
            vm.IsDirectoryChooser = true;
            bool? result = vm.Show();
            if (result == true) 
            {
                outPath = vm.SelectedFolder.Path;
                OutputLabel.Content = String.Format("Сохранять результат в: {0}", outPath.ToString());
                IsOutputPathSelected = true;
                CheckIsRunEnabled();
            }
        }

        private void readXls(String fileName)
        {
            var file = new FileInfo(fileName);
            using (var stream = new FileStream(fileName, FileMode.Open))
            {
                IExcelDataReader reader = null;
                if (file.Extension == ".xls")
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (file.Extension == ".xlsx")
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    Console.WriteLine("fail");

                reader.IsFirstRowAsColumnNames = true;
                ds = reader.AsDataSet();

            }

        }

        private IList<string> GetTablenames(DataTableCollection tables)
        {
            var tableList = new List<string>();
            foreach (var table in tables)
            {
                tableList.Add(table.ToString());
            }
            return tableList;
        }


        private void Process(string SheetName, UserDoc userDoc)
        {
            using (var stream = new MemoryStream())
            {
                using (var fileStream = new FileStream(userDoc.Path, FileMode.Open))
                {
                    fileStream.CopyTo(stream);
                }

                var template = new DocTemplate(stream);
                var cols = ds.Tables[SheetName].Columns;

                System.Threading.Tasks.Parallel.ForEach(
                    Enumerable.Range(0, ds.Tables[SheetName].Rows.Count),
                    i =>
                    {
                        var bm = new BindMap(ds.Tables[SheetName].Rows[i], cols);
                        template.CreateDocument(System.IO.Path.Combine(outPath, (i + 1).ToString() + ".docx"), bm);
                    }
                );
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            readXls(inputExcelFile.Path);
            var tablenames = GetTablenames(ds.Tables);
            Process(tablenames[0], inputDocFiles[0]);
        }
    }

}
