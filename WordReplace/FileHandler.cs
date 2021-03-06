﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;


namespace WordReplace
{
    class FileHandler
    {
         public static UserInputFile Handle(string path)
        {
             string ext = Path.GetExtension(path);
             switch (ext)
             {
                 case ".docx":
                    return new UserDoc(path);
                 case ".xlsx":
                    return new UserExcel(path);
                 default:
                    throw new ArgumentException();
             }
        }
    }


    public class UserInputFile
    {
        public string Name { get; set; }
        public string Path { get; set; }

        public UserInputFile(string path)
        {
            this.Name = System.IO.Path.GetFileName(path);
            this.Path = path;
        }
    }

    public class UserExcel : UserInputFile, IDisposable
    {
        public DataSet ds = new DataSet();
        private bool fake = false;

        public UserExcel(string path, bool fake = false) : base(path)
        {
            if (!fake)
            {
                getDataTableFromExcel();
            }
            this.fake = fake;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                ds.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void getDataTableFromExcel()
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(Path))
                {
                    pck.Load(stream);
                }
                string cellText;
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    DataTable tbl = ds.Tables.Add(ws.Name);
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        try
                        {
                            tbl.Columns.Add(firstRowCell.Text);
                        }
                        catch (DuplicateNameException)
                        {
                            //TODO: refactor
                            throw;
                        }

                    }
                    Enumerable.Range(tbl.Columns.Count, ws.Dimension.End.Column - tbl.Columns.Count)
                        .ToList()
                        .ForEach(i => tbl.Columns.Add("Column" + i));
                    for (var rowNum = 2; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        var row = tbl.NewRow();
                        foreach (var cell in wsRow)
                        {
                            
                            try
                            {
                                cellText = cell.Text;
                            }
                            catch (ArgumentOutOfRangeException exc)
                            {
                                cellText = cell.Value.ToString();
                                Console.WriteLine(String.Format("Can't get Text of cell {0} ({1})", cell.Address, exc.Message));
                            }
                            row[cell.Start.Column - 1] = cellText;
                        }
                        tbl.Rows.Add(row);
                    }
                    break;
                }
            }
        }
    }

    public class UserDoc : UserInputFile
    {
        private static Regex RGX_BAD_CHAR = new Regex(@"[^\w\.\s\-]");
        private static Regex RGX_WHITESPACE = new Regex(@"\s+");

        public string NoExtName { get; set; }

        public UserDoc(string path) : base(path)
        {
            this.NoExtName = System.IO.Path.GetFileNameWithoutExtension(Path);
        }

        internal List<Tuple<BindMap, string>> MapOutputNames(IEnumerable<BindMap> bms, string outPath, List<string> columnNames)
        {
            // we don't handle case when user provided two doc files
            // with same name; in this case resulting files will be rewrited
            return bms
                .Select(bm => new Tuple<BindMap, string>(bm, buildUniqOutputPath(bm, columnNames)))
                .GroupBy(row => row.Item2)
                .SelectMany(group =>
                {
                    return Enumerable.Range(0, group.Count())
                        .Zip(group, (i, row) =>
                        {
                            List<string> parts = new List<string>();
                            if (!String.IsNullOrWhiteSpace(row.Item2))
                                parts.Add(row.Item2);
                            parts.Add(NoExtName);
                            if (i > 0)
                                parts.Add(i.ToString());
                            string outName = String.Join("-", parts) + ".docx";
                            outName = System.IO.Path.Combine(outPath, outName);
                            return new Tuple<BindMap, string>(row.Item1, outName);
                        });
                }).ToList();
        }

        private string buildUniqOutputPath(BindMap bm, List<string> columnNames)
        {
            List<string> names = columnNames
                .Select(n => bm.Get(n, ""))
                .Where(n => !String.IsNullOrWhiteSpace(n))
                .ToList();

            string result = String.Join("-", names);
            result = RGX_WHITESPACE.Replace(result, " ");
            result = RGX_BAD_CHAR.Replace(result, "");
            return result;
        }
    }
}
