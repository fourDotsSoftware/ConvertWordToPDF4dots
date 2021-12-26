using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ConvertWordToPDF4dots
{ 
    class ArgsHelper
    {        
        public static bool ExamineArgs(string[] args)
        {
            if (args.Length == 0) return true;
                        
            Module.args = args;

            try
            {
                if (args[0].ToLower().Trim().StartsWith("-tempfile:"))
                {                                       
                    string tempfile = GetParameter(args[0]);

                    //MessageBox.Show(tempfile);

                    using (StreamReader sr = new StreamReader(tempfile, Encoding.Unicode))
                    {
                        string scont = sr.ReadToEnd();

                        //args = scont.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        args = SplitArguments(scont);
                        Module.args = args;

                        // MessageBox.Show(scont);
                    }
                }
                else if (args.Length>0 && (Module.args.Length==1 && (System.IO.File.Exists(Module.args[0]) || System.IO.Directory.Exists(Module.args[0]))))
                {

                }
                else
                {
                    Module.IsCommandLine = true;

                    //System.Windows.Forms.MessageBox.Show("0");

                    //1frmMain f=new frmMain();  

                    //frmMain.Instance.SetupOnLoad();

                    for (int k = 0; k < Module.args.Length; k++)
                    {
                        if (System.IO.File.Exists(Module.args[k]))
                        {
                            frmMain.Instance.AddFile(Module.args[k]);                            
                        }
                        else if (System.IO.Directory.Exists(Module.args[k]))
                        {
                            frmMain.Instance.SilentAdd = true;

                            frmMain.Instance.AddFolder(Module.args[k]);                            
                        }                        
                        else if (Module.args[k].ToLower().StartsWith("/outputfolder:") ||
        Module.args[k].ToLower().StartsWith("-outputfolder:"))
                        {
                            string outfolder = GetParameter(Module.args[k]);

                            frmMain.Instance.cmbOutputDir.Items.Add(outfolder);
                            frmMain.Instance.cmbOutputDir.SelectedIndex = frmMain.Instance.cmbOutputDir.Items.Count - 1;

                            //frmMain.Instance.cmbOutputDir.Text = outfolder;
                        }
                        else if (Module.args[k].ToLower().StartsWith("/importtext:") ||
        Module.args[k].ToLower().StartsWith("-importtext:"))
                        {
                            string lf = GetParameter(Module.args[k]);

                            frmMain.Instance.ImportList(lf);
                        }
                        else if (Module.args[k].ToLower().StartsWith("/importexcel:") ||
        Module.args[k].ToLower().StartsWith("-importexcel:"))
                        {
                            string lf = GetParameter(Module.args[k]);

                            ExcelImporter xl = new ExcelImporter();
                            xl.ImportListExcel(lf);
                        }
                        else if (Module.args[k].ToLower() == "/h" ||
                        Module.args[k].ToLower() == "-h" ||
                        Module.args[k].ToLower() == "-?" ||
                        Module.args[k].ToLower() == "/?")
                        {
                            ShowCommandUsage();
                            Environment.Exit(1);
                            return true;
                        }
                    }                                      
                }
            }
            catch (Exception ex)
            {
                Module.ShowError("Error could not parse Arguments !", ex.ToString());
                return false;
            }

            return true;
        }

        private static string GetParameter(string arg)
        {
            int spos = arg.IndexOf(":");
            if (spos == arg.Length - 1) return "";
            else
            {
                string str=arg.Substring(spos + 1);

                if ((str.StartsWith("\"") && str.EndsWith("\"")) ||
                    (str.StartsWith("'") && str.EndsWith("'")))
                {
                    if (str.Length > 2)
                    {
                        str = str.Substring(1, str.Length - 2);
                    }
                    else
                    {
                        str = "";
                    }
                }

                return str;
            }
        }

        public static string[] SplitArguments(string commandLine)
        {
            char[] parmChars = commandLine.ToCharArray();
            bool inSingleQuote = false;
            bool inDoubleQuote = false;
            for (int index = 0; index < parmChars.Length; index++)
            {
                if (parmChars[index] == '"' && !inSingleQuote)
                {
                    inDoubleQuote = !inDoubleQuote;
                    parmChars[index] = '\n';
                }
                if (parmChars[index] == '\'' && !inDoubleQuote)
                {
                    inSingleQuote = !inSingleQuote;
                    parmChars[index] = '\n';
                }
                if (!inSingleQuote && !inDoubleQuote && parmChars[index] == ' ')
                    parmChars[index] = '\n';
            }
            return (new string(parmChars)).Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public static void ShowCommandUsage()
        {
            string msg = "Batch convert Word to PDF. PPTX to PDF. PPT to PDF.\n\n" +
            "ConvertWordToPDF4dots.exe [[file|directory]]\n" +
            "[/outputfolder:OUTPUT_FOLDER_VALUE]\n" +
            "[/importtext:IMPORT_TEXT_FILE]\n"+
            "[/importexcel:IMPORT_EXCEL_FILE]\n" +
            "[/?]\n\n\n" +
            "file : one or more files to be processed.\n" +
            "directory : one or more directories containing files to be processed.\n" +
            "outputfolder: Output folder value (if different than the folder of the first file)\n" +
            "importtext : import list from Text file\n"+
            "importexcel : import list from Excel file\n" +
            "/? : show help\n\n\n" +
            "Example :\n" +
            "ConvertWordToPDF4dots.exe \"c:\\documents\\presentation.pptx\"\n\n" +
            "ConvertWordToPDF4dots.exe \"c:\\documents\\presentations\"\n\n"+
            "ConvertWordToPDF4dots.exe /importtext:\"c:\\documents\\list.txt\"\n\n" +
            "ConvertWordToPDF4dots.exe /importexcel:\"c:\\documents\\list.xlsx\"\n\n";

            Module.ShowMessage(msg);

            Environment.Exit(0);
        }

        public static bool IsFromFolderWatcher
        {
            get
            {                
                // new
                if (Module.args.Length > 0 && Module.args[0].ToLower().Trim() == "/cmdfw")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public static bool IsFromWindowsExplorer
        {
            get
            {
                if (Module.IsFromWindowsExplorer) return true;

                // new
                if (Module.args.Length > 0 && (Module.args[0].ToLower().Trim().Contains("-tempfile:")
                    || (Module.args.Length==1 && (System.IO.File.Exists(Module.args[0]) || System.IO.Directory.Exists(Module.args[0])))))
                {
                    Module.IsFromWindowsExplorer = true;
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        public static bool IsFromCommandLine
        {
            get
            {
                if (Module.args == null || Module.args.Length == 0)
                {
                    return false;
                }

                if (ArgsHelper.IsFromWindowsExplorer)
                {
                    Module.IsCommandLine = false;
                    return false;
                }
                else
                {
                    Module.IsCommandLine = true;
                    return true;
                }
            }
        }

        /*
        public static bool IsFromWindowsExplorer()
        {
            if (Module.args == null || Module.args.Length == 0)
            {
                return false;
            }

            for (int k = 0; k < Module.args.Length; k++)
            {
                if (Module.args[k] == "-visual")
                {
                    Module.IsFromWindowsExplorer = true;
                    return true;
                }
            }

            Module.IsFromWindowsExplorer = false;
            return false;
        }
        */

        public static void ExecuteCommandLine()
        {
            string err = "";
            bool finished = false;

            try
            {
                /*
                if (Module.CmdLogFile != string.Empty)
                {
                    try
                    {
                        Module.CmdLogFileWriter = new StreamWriter(Module.CmdLogFile, true);
                        Module.CmdLogFileWriter.AutoFlush = true;
                        Module.CmdLogFileWriter.WriteLine("[" + DateTime.Now.ToString() + "] Started compressing PDF files !");
                    }
                    catch (Exception exl)
                    {
                        Module.ShowMessage("Error could not start log writer !");
                        ShowCommandUsage();
                        Environment.Exit(0);
                        return;
                    }
                }                

                if (Module.CmdImportListFile != string.Empty)
                {
                    frmMain.Instance.ImportList(Module.CmdImportListFile);

                    err += frmMain.Instance.SilentAddErr;

                }
                */

                if (frmMain.Instance.dt.Rows.Count == 0)
                {
                    Module.ShowMessage("Please specify Files !");
                    ShowCommandUsage();
                    Environment.Exit(0);
                    return;
                }

                Console.Clear();

                Module.ShowMessage("Please wait...\nPress ^C (Control + C) to cancel operation.");

                Console.CancelKeyPress += Console_CancelKeyPress;

                bwMsg.DoWork += BwMsg_DoWork;
                bwMsg.WorkerReportsProgress = true;
                bwMsg.WorkerSupportsCancellation = true;
                bwMsg.ProgressChanged += BwMsg_ProgressChanged;
                bwMsg.RunWorkerAsync();

                frmMain.Instance.tsbConvertWordToPDF_Click(null, null);

                bwMsg.CancelAsync();

                Application.Exit();                             
            }
            finally
            {
                
            }
            Environment.Exit(0);
        }

        private static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            Console.WriteLine("Operation cancelled !");

            try
            {
                frmMain.Instance.OperationStopped = true;
            }
            catch { }

            Environment.Exit(1);
        }

        private static void BwMsg_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            Console.Write(".");
        }

        private static void BwMsg_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            while (true)
            {
                if (bwMsg.CancellationPending)
                {
                    return;
                }
                else
                {
                    bwMsg.ReportProgress(0);
                    System.Threading.Thread.Sleep(1500);
                }
            }
        }

        public static System.ComponentModel.BackgroundWorker bwMsg = new System.ComponentModel.BackgroundWorker();





    }

    public class ReadListsResult
    {
        public bool Success = true;
        public string err = "";
    }
}
