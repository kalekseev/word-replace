using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


namespace WordReplace
{
    public partial class MainWindow : Window, IDisposable
    {
        private ObservableCollection<UserInputFile> inputDocFiles = new ObservableCollection<UserInputFile>();
        private UserExcel inputExcelFile;
        private string outPath = Properties.Settings.Default.outPath;
        private ProgressWindow pWin = new ProgressWindow();
        private List<string> columnNames = new List<string>();
        private CancellationTokenSource cts;


        public MainWindow()
        {
            InitializeComponent();
            //TODO: fix this
            inputDocFiles.Add(new UserExcel("D:\\No excel file provided", true));
            FileListBox.ItemsSource = inputDocFiles;
        }

        private void OnLoaded(object sender, EventArgs eventArgs)
        {
            pWin.Owner = this;
            pWin.Closing += new CancelEventHandler(OnProgressClosing);
            OutputLabel.Content = outPath;
        }

        private void OnProgressClosing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            if (cts != null)
            {
                cts.Cancel();
            }
        }

        private void ProcessFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            pWin.Hide();
            MainGrid.IsEnabled = true;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                pWin.Close();
                if (cts != null)
                {
                    cts.Cancel();
                    cts.Dispose();
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
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
                    UserInputFile newFile = FileHandler.Handle(f);
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
                            inputDocFiles[0] = newFile;
                            var cols = inputExcelFile.ds.Tables[0].Columns;
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
                    }
                    CheckIsRunEnabled();   
                }
                catch (ArgumentException)
                {
                }
                catch (IOException exc)
                {
                    MessageBoxResult result = MessageBox.Show(exc.Message, "Error");
                }
                catch (DuplicateNameException exc)
                {
                    MessageBoxResult result = MessageBox.Show(exc.Message, "Error");
                }
            }
        }

        private void CheckIsRunEnabled()
        {
            RunButton.IsEnabled = (inputExcelFile != null && inputDocFiles.Count > 1 && outPath != null);
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
                OutputLabel.Content = outPath;
                Properties.Settings.Default.outPath = outPath;
                Properties.Settings.Default.Save();
                CheckIsRunEnabled();
            }
        }

        private void Process(object sender, EventArgs e)
        {
            var tbl = inputExcelFile.ds.Tables[0];
            var totalRows = tbl.Rows.Count * (inputDocFiles.Count - 1);
            foreach (UserDoc userDoc in inputDocFiles.Skip(1))
            {
                using (var stream = new MemoryStream())
                {
                    using (var fileStream = new FileStream(userDoc.Path, FileMode.Open))
                        fileStream.CopyTo(stream);

                    var template = new DocTemplate(stream);
                    var po = new ParallelOptions();
                    cts = new CancellationTokenSource();
                    po.CancellationToken = cts.Token;
                    IEnumerable<BindMap> rows = Enumerable.Range(0, tbl.Rows.Count)
                        .Select(i => new BindMap(tbl.Rows[i], tbl.Columns))
                        .Where(bm => !bm.isEmpty());
                    List<Tuple<BindMap, string>> rowsWithName = userDoc.MapOutputNames(rows, outPath, columnNames);

                    System.Threading.Tasks.Parallel.ForEach(
                        rowsWithName,
                        po,
                        row =>
                        {
                            template.CreateDocument(row.Item2, row.Item1);
                            (sender as BackgroundWorker).ReportProgress(totalRows);
                        }
                    );
                }
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainGrid.IsEnabled = false;
            pWin.Reset();
            pWin.Show();
            columnNames = (new List<ComboBox> { FileNameSelect1, FileNameSelect2, FileNameSelect3 })
                .Where(cb => cb.SelectedIndex > 0)
                .Select(cb => cb.SelectedItem as string)
                .ToList();
            var worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += Process;
            worker.ProgressChanged += pWin.Update;
            worker.RunWorkerAsync();
            worker.RunWorkerCompleted += ProcessFinished;
        }

        private void Clear_SelectedDocuments(object sender, RoutedEventArgs e)
        {
            inputDocFiles.Clear();
            inputDocFiles.Add(new UserExcel("D:\\No excel file provided", true));
            FileNameSelect1.ItemsSource = null;
            FileNameSelect2.ItemsSource = null;
            FileNameSelect3.ItemsSource = null;
            CheckIsRunEnabled();
        }
    } 
}
