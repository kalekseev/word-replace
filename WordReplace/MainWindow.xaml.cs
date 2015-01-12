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
using System.Text.RegularExpressions;
using OfficeOpenXml;
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
                            readXls(inputExcelFile.Path);
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
                OutputLabel.Content = outPath.ToString();
                IsOutputPathSelected = true;
                CheckIsRunEnabled();
            }
        }

        public static DataSet getDataTableFromExcel(string path)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var dSet = new DataSet();
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    DataTable tbl = dSet.Tables.Add(ws.Name);
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(firstRowCell.Text);
                    }
                    for (var rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        var row = tbl.NewRow();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                        tbl.Rows.Add(row);
                    }
                    break;
                }
                return dSet;
            }
        }

        private void readXls(String fileName)
        {
            var file = new FileInfo(fileName);
            ds = getDataTableFromExcel(fileName);

            var cols = ds.Tables[0].Columns;
            List<string> colNames = new List<string>();
            colNames.Add("---");
            foreach (DataColumn col in cols)
                colNames.Add(col.ColumnName);
            FileNameSelect1.ItemsSource = colNames;
            FileNameSelect1.SelectedIndex = 0;
            FileNameSelect2.ItemsSource = colNames;
            FileNameSelect2.SelectedIndex = 0;
            FileNameSelect3.ItemsSource = colNames;
            FileNameSelect3.SelectedIndex = 0;
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


        private void Process(UserDoc userDoc)
        {
            using (var stream = new MemoryStream())
            {
                using (var fileStream = new FileStream(userDoc.Path, FileMode.Open))
                {
                    fileStream.CopyTo(stream);
                }

                var template = new DocTemplate(stream);
                var tbl = ds.Tables[0];
                var cols = tbl.Columns;
                List<ComboBoxInfo> cbs = new List<ComboBoxInfo>();
                cbs.Add(new ComboBoxInfo { Index = FileNameSelect1.SelectedIndex, Name = (FileNameSelect1.SelectedItem as string) });
                cbs.Add(new ComboBoxInfo { Index = FileNameSelect2.SelectedIndex, Name = (FileNameSelect1.SelectedItem as string) });
                cbs.Add(new ComboBoxInfo { Index = FileNameSelect3.SelectedIndex, Name = (FileNameSelect1.SelectedItem as string) });

                System.Threading.Tasks.Parallel.ForEach(
                    Enumerable.Range(0, tbl.Rows.Count),
                    i =>
                    {
                        var bm = new BindMap(tbl.Rows[i], cols);
                        template.CreateDocument(buildOutputPath(cbs, bm, i), bm);
                    }
                );
            }

        }

        private void addField(List<string> names, ComboBoxInfo comboBox, BindMap bm)
        {
            if (comboBox.Index > 0)
            {
                var key = comboBox.Name as string;
                try
                {
                    names.Add(bm.Get(key));
                }
                catch { }
            }
        }

        private string buildOutputPath(List<ComboBoxInfo> cbs, BindMap bm, int i)
        {
            List<string> names = new List<string>();
            foreach (ComboBoxInfo cb in cbs)
                addField(names, cb, bm);


            Regex rgxBadChar = new Regex(@"[^\w-]");
            Regex rgxWhiteSpace = new Regex(@"\s+");
            string result = String.Join("-", names);
            result = rgxWhiteSpace.Replace(result, "_");
            result = rgxBadChar.Replace(result, "");
            result += "-" + (i + 1).ToString() + ".docx";
            return System.IO.Path.Combine(outPath, result);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            Process(inputDocFiles[0]);
        }
    }

    public class ComboBoxInfo
    {
        public int Index { get; set; }
        public string Name { get; set; }
    }

}
