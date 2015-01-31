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
using System.ComponentModel;
using System.Threading;

namespace WordReplace
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<UserInputFile> inputDocFiles = new ObservableCollection<UserInputFile>();
        private UserExcel inputExcelFile;
        private DataSet ds;
        private string outPath;
        private FileHandler fh = new FileHandler();
        private Boolean IsOutputPathSelected = false;
        private ProgressWindow pWin = new ProgressWindow();
        private List<ComboBoxInfo> cbs = new List<ComboBoxInfo>();
        private CancellationTokenSource cts;


        public MainWindow()
        {
            InitializeComponent();
            inputDocFiles.Add(new UserExcel("D:\\No excel file provided"));
            FileListBox.ItemsSource = inputDocFiles;
        }

        private void OnLoaded(object sender, EventArgs eventArgs)
        {
            pWin.Owner = this;
            pWin.Closing += new CancelEventHandler(OnProgressClosing);
        }

        private void OnClosing(object sender, EventArgs eventArgs)
        {
            pWin.Close();
        }

        private void OnProgressClosing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            if (cts != null)
            {
                cts.Cancel();
            }
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
                        inputDocFiles.Add(newFile);
                    }
                    else
                    {
                        UserExcel newExcel = newFile as UserExcel;
                        if (newExcel != null)
                        {
                            inputExcelFile = newExcel;
                            readXls(inputExcelFile.Path);
                            inputDocFiles[0] = newFile;
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
            if (inputExcelFile != null && inputDocFiles.Count > 1 && IsOutputPathSelected)
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
            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderDialog.Description = "Select output path";
            folderDialog.SelectedPath = outPath;
            folderDialog.ShowNewFolderButton = true;
            System.Windows.Forms.DialogResult result = folderDialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                outPath = folderDialog.SelectedPath;
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


        private void Process(object sender, EventArgs e)
        {
            UserDoc userDoc = inputDocFiles[1] as UserDoc;
            using (var stream = new MemoryStream())
            {
                using (var fileStream = new FileStream(userDoc.Path, FileMode.Open))
                {
                    fileStream.CopyTo(stream);
                }

                var template = new DocTemplate(stream);
                var tbl = ds.Tables[0];
                var cols = tbl.Columns;

                ParallelOptions po = new ParallelOptions();
                cts = new CancellationTokenSource();
                po.CancellationToken = cts.Token;
                List<Tuple<BindMap, string>> rows = Enumerable.Range(0, tbl.Rows.Count)
                    .Select(i =>
                    {
                        var bm = new BindMap(tbl.Rows[i], cols);
                        return new Tuple<BindMap, string>(bm, buildOutputPath(bm));
                    })
                    .GroupBy(row => row.Item2)
                    .SelectMany(group =>
                    {
                        return Enumerable.Range(0, group.Count())
                            .Zip(group, (i, row) => {
                                string outName = userDoc.NoExtName;
                                if (!String.IsNullOrWhiteSpace(row.Item2))
                                    outName += "-" + row.Item2;
                                if (i > 0)
                                    outName += "-" + i.ToString();
                                outName = outName + ".docx";
                                outName = System.IO.Path.Combine(outPath, outName);
                                return new Tuple<BindMap, string>(row.Item1, outName);
                            });
                    }).ToList();
                System.Threading.Tasks.Parallel.ForEach(
                    rows,
                    po,
                    row =>
                    {
                        template.CreateDocument(row.Item2, row.Item1);
                        //(sender as BackgroundWorker).ReportProgress((100 / tbl.Rows.Count) * (i + 1));
                    }
                );
            }

        }

        private string buildOutputPath(BindMap bm)
        {
            List<string> names = cbs
                .Where(cb => cb.Index > 0)
                .Select(cb => {
                    var name = cb.Name as string;
                    return bm.Get(name, "");
                })
                .Where(name => !String.IsNullOrWhiteSpace(name))
                .ToList();

            Regex rgxBadChar = new Regex(@"[^\w-]");
            Regex rgxWhiteSpace = new Regex(@"\s+");
            string result = String.Join("-", names);
            result = rgxWhiteSpace.Replace(result, "_");
            result = rgxBadChar.Replace(result, "");
            return result;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainGrid.IsEnabled = false;
            pWin.Reset();
            pWin.Show();
            cbs.Clear();
            cbs.Add(new ComboBoxInfo { Index = FileNameSelect1.SelectedIndex, Name = (FileNameSelect1.SelectedItem as string) });
            cbs.Add(new ComboBoxInfo { Index = FileNameSelect2.SelectedIndex, Name = (FileNameSelect2.SelectedItem as string) });
            cbs.Add(new ComboBoxInfo { Index = FileNameSelect3.SelectedIndex, Name = (FileNameSelect3.SelectedItem as string) });
            BackgroundWorker worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += Process;
            worker.ProgressChanged += pWin.Update;
            worker.RunWorkerAsync();
            worker.RunWorkerCompleted += ProcessFinished;
        }

        private void ProcessFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            pWin.Hide();
            MainGrid.IsEnabled = true;
        }
    } 

    public class ComboBoxInfo
    {
        public int Index { get; set; }
        public string Name { get; set; }
    }
}
