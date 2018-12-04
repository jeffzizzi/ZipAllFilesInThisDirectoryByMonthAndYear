using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;

namespace ZipAllFilesInThisDirectoryByMonthAndYear
{
    class Program
    {

        private static readonly log4net.ILog log =
    log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static string thisDir = "";
        static bool deleteFile = false;
        static string[] FilesToIgnore = null;
        static void Main(string[] args)
        {
            bool.TryParse(ConfigurationManager.AppSettings["DeleteFilesAfterZipping"], out deleteFile);

            FilesToIgnore = ConfigurationManager.AppSettings["FilesToIgnoreInCsv"].Split(',');

            if (deleteFile == false)
                log.Info("Starting app with Delete Files turned off.");

            thisDir = ConfigurationManager.AppSettings["StartingDirectory"];
            log.Info($"Starting zipping of all files in this directory: {thisDir}");
            //start it off
            ProcessDirectory(thisDir);

            log.Info("Completed");
        }

        static void ProcessDirectory(string dirPath)
        {
            Console.Write($"Processing directory {dirPath}");
            string[] allFilesAndDirectories = Directory.GetFiles(dirPath, "*.*", SearchOption.AllDirectories);
            foreach (string item in allFilesAndDirectories)
            {
                FileAttributes attr = File.GetAttributes(item);

                if (attr.HasFlag(FileAttributes.Directory))
                    ProcessDirectory(item);
                else
                    ProcessFile(item);
            }
        }


        static void ProcessFile(string filePath)
        {
            log.Info($"Processing file {filePath}");
            if (filePath.ToUpper().Contains(Path.GetFileNameWithoutExtension(System.AppDomain.CurrentDomain.FriendlyName.ToUpper())) == false &&
                isIgnoreFile(filePath) == false)
            {
                //Read the date last modified
                FileInfo fi = new FileInfo(filePath);
                string ZipFile = Path.Combine(fi.Directory.FullName, string.Format("{0}_{1}.7z", fi.Directory.Name, fi.LastWriteTime.ToString("yyyyMM")));
                log.Debug($"Calling AddFileToZip with {ZipFile} and {filePath}");
                AddFileToZip(filePath, ZipFile);
            }

        }

        static bool isIgnoreFile(string file)
        {
            bool rtnVal = false;
            foreach (string searchTerm in FilesToIgnore)
            {
                if (file.ToUpper().Contains(searchTerm.ToUpper()))
                    return true;
            }
            return rtnVal;
        }

        static void AddFileToZip(string filePath, string zipFile)
        {
            string programUsedToZip = @"C:\Program Files\7-Zip\7z.exe";
            string arguments = string.Format("u -mx9 -t7z -m0=lzma2 \"{0}\" \"{1}\"", zipFile, filePath);
            Process proc = new Process();
            proc.StartInfo.FileName = programUsedToZip;
            proc.StartInfo.Arguments = arguments;
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.Start();
            proc.WaitForExit();
            if (proc.ExitCode == 0)
            {
                if (deleteFile == true)
                {
                    File.Delete(filePath);
                }
            }

        }
    }
}
