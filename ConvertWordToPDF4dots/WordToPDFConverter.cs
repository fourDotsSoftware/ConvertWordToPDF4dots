using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace ConvertWordToPDF4dots
{
    public class WordToPDFConverter
    {        
        public string err = "";

        public bool ConvertToPDF(string filepath,string outfilepath)
        {
            err = "";
            
            object oDocuments = null;
            object doc = null;

            try
            {
                OfficeHelper.CreateWordApplication();

                oDocuments = OfficeHelper.WordApp.GetType().InvokeMember("Documents", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);                

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);

                OfficeHelper.WordApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.WordApp, null);
                */

                System.Threading.Thread.Sleep(200);

                /*
                string fp=System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(filepath),
                    System.IO.Path.GetFileNameWithoutExtension(filepath)+".pdf"
                    );                
                */

                doc.GetType().InvokeMember("ExportAsFixedFormat", BindingFlags.InvokeMethod, null, doc, new object[] { outfilepath, 17 });

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Word to PDF") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }                
    }
}