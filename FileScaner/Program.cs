using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FileScaner
{
    class Program
    {
        static void Main(string[] args)
        {
            string strPath = "G:\\美化大师背景模板";
            string strSavePath = @"G:\App\Template\";
            string[] strFiles = Directory.GetFiles(strPath, "*", SearchOption.AllDirectories);
            strFiles = strFiles.Where(f => f.EndsWith(".pptx") && f.IndexOf("KPBG") > 0).ToArray();
            int i = 0;
            int count = strFiles.Length;
            foreach (string file in strFiles)
            {
                string strNewPath = strSavePath + Path.GetFileName(file);
                if(File.Exists(strNewPath))
                    continue;
                File.Move(file, strSavePath + Path.GetFileName(file));
                Console.WriteLine((++i) + "/" + count);
            }
        }
    }
}
