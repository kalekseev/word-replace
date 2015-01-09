using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.IO;

namespace WordReplace
{
    class FileHandler
    {
         public UserInputFile Handle(string path)
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

    public class UserExcel : UserInputFile
    {
        public UserExcel(string path)
            : base(path)
        {

        }
    }

    public class UserDoc : UserInputFile
    {
        public UserDoc(string path)
            : base(path)
        {

        }
    }
}
