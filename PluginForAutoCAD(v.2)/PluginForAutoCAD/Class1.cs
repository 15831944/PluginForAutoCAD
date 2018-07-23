using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Acad = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using Autodesk.AutoCAD.Publishing;
using Autodesk.AutoCAD.PlottingServices;
using System.Text;

namespace AutocadPlugin
{
    public class MultiSheetsPdf
    {
        private string dwgFile, pdfFile, dsdFile, outputDir;
        private int sheetNum;
        IEnumerable<Layout> layouts;

        private const string LOG = "publish.log";

        public MultiSheetsPdf(string pdfFile, IEnumerable<Layout> layouts)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            this.dwgFile = db.Filename;
            this.pdfFile = pdfFile;
            this.outputDir = Path.GetDirectoryName(this.pdfFile);
            this.dsdFile = Path.ChangeExtension(this.pdfFile, "dsd");
            this.layouts = layouts;
        }

        public void Publish()
        {
            if (TryCreateDSD())
            {
                Publisher publisher = Acad.Publisher;
                PlotProgressDialog plotDlg = new PlotProgressDialog(false, this.sheetNum, true);
                plotDlg.IsVisible = false;
                publisher.PublishDsd(this.dsdFile, plotDlg);
                plotDlg.Destroy();
                File.Delete(this.dsdFile);
            }
        }

        private bool TryCreateDSD()
        {
            using (DsdData dsd = new DsdData())
            using (DsdEntryCollection dsdEntries = CreateDsdEntryCollection(this.layouts))
            {
                if (dsdEntries == null || dsdEntries.Count <= 0) return false;

                if (!Directory.Exists(this.outputDir))
                    Directory.CreateDirectory(this.outputDir);

                this.sheetNum = dsdEntries.Count;

                dsd.SetDsdEntryCollection(dsdEntries);

                dsd.SetUnrecognizedData("PwdProtectPublishedDWF", "FALSE");
                dsd.SetUnrecognizedData("PromptForPwd", "FALSE");
                dsd.SheetType = SheetType.MultiDwf;
                dsd.NoOfCopies = 1;
                dsd.DestinationName = this.pdfFile;
                dsd.IsHomogeneous = false;
                dsd.LogFilePath = Path.Combine(this.outputDir, LOG);

                PostProcessDSD(dsd);

                return true;
            }
        }

        private DsdEntryCollection CreateDsdEntryCollection(IEnumerable<Layout> layouts)
        {
            DsdEntryCollection entries = new DsdEntryCollection();

            foreach (Layout layout in layouts)
            {
                DsdEntry dsdEntry = new DsdEntry();
                dsdEntry.DwgName = this.dwgFile;
                dsdEntry.Layout = layout.LayoutName;
                dsdEntry.Title = Path.GetFileNameWithoutExtension(this.dwgFile) + "-" + layout.LayoutName;
                dsdEntry.Nps = layout.TabOrder.ToString();
                entries.Add(dsdEntry);
            }
            return entries;
        }

        private void PostProcessDSD(DsdData dsd)
        {
            string str, newStr;
            string tmpFile = Path.Combine(this.outputDir, "temp.dsd");

            dsd.WriteDsd(tmpFile);

            using (StreamReader reader = new StreamReader(tmpFile, Encoding.Default))
            using (StreamWriter writer = new StreamWriter(this.dsdFile, false, Encoding.Default))
            {
                while (!reader.EndOfStream)
                {
                    str = reader.ReadLine();
                    if (str.Contains("Has3DDWF"))
                    {
                        newStr = "Has3DDWF=0";
                    }
                    else if (str.Contains("OriginalSheetPath"))
                    {
                        newStr = "OriginalSheetPath=" + this.dwgFile;
                    }
                    else if (str.Contains("Type"))
                    {
                        newStr = "Type=6";
                    }
                    else if (str.Contains("OUT"))
                    {
                        newStr = "OUT=" + this.outputDir;
                    }
                    else if (str.Contains("IncludeLayer"))
                    {
                        newStr = "IncludeLayer=TRUE";
                    }
                    else if (str.Contains("PromptForDwfName"))
                    {
                        newStr = "PromptForDwfName=FALSE";
                    }
                    else if (str.Contains("LogFilePath"))
                    {
                        newStr = "LogFilePath=" + Path.Combine(this.outputDir, LOG);
                    }
                    else
                    {
                        newStr = str;
                    }
                    writer.WriteLine(newStr);
                }
            }
            File.Delete(tmpFile);
        }
    }
    public class test : IExtensionApplication
    {
        
        [CommandMethod("hello")]
        public void Helloworld()
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Привет из Autocad плагина");
        }

        public void Initialize()
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Инициализация плагина.." + Environment.NewLine);

        }

        public void Terminate()
        {

        }
      /*  [CommandMethod("SaveAsPDF")]
        public void SaveAsPDF()
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Процесс экспорта чертежей в формат PDF начинается...");

        }*/
        [CommandMethod("HelpSaveAsPDF")]
        public void Help()
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Для запуска пакетного сохранения листов чертежей в формате .pdf" + Environment.NewLine +
                "наберите в командной строке команду в следующем формате:" + Environment.NewLine +
                "(saveaspdf " + '"' +"диск:/папка1/.../папка с чертежами" +'"' + " "+ '"' + "true/false" + '"' + ")");
        }
    }
    public class test1
    {
        [LispFunction("SaveAsPDF")]
        public static void SaveAsPDF(ResultBuffer rbArgs)
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Процесс экспорта чертежей в формат PDF начинается...");
            // Получение директории и параметров поиска
            if (rbArgs != null)
            {
                string str = "";
                bool subDirs = false;
                foreach (TypedValue rb in rbArgs)
                {
                    if (rb.TypeCode == (int)Autodesk.AutoCAD.Runtime.LispDataType.Text)
                    {
                        if (rb.Value.ToString() == "true")
                            subDirs = true;
                        else if (rb.Value.ToString() == "false")
                            subDirs = false;
                        else
                            str += rb.Value.ToString();
                    }
                }
                Acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Путь к чертежам: " + str);
                // Получение списка файлов с учетом заданного пути и параметров поиска
                try
                {
                    str = str.Replace('/', '\\');
                    System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(str);
                    System.IO.FileInfo[] dirs = rootDir.GetFiles("*.dwg", subDirs ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
                    //Список листов найденных чертежей
                    foreach (FileInfo dir in dirs)
                    {
                        editor.WriteMessage(Environment.NewLine + dir);
                        using (Database dwgDB = new Database(false, true))
                        {
                            dwgDB.ReadDwgFile(dir.FullName, System.IO.FileShare.ReadWrite, false, "");
                            Acad.SetSystemVariable("BACKGROUNDPLOT", 1);
                            //Далее манипуляции с базой данных чертежа
                            using (Transaction acTrans = dwgDB.TransactionManager.StartTransaction())
                            {
                                List<Layout> layouts = new List<Layout>();
                                DBDictionary layoutDict = (DBDictionary)acTrans.GetObject(dwgDB.LayoutDictionaryId, OpenMode.ForRead);
                                editor.WriteMessage("Листы чертежа " + dir + ":");
                                foreach (DBDictionaryEntry id in layoutDict)
                                {
                                    if (id.Key != "Model")
                                    {
                                        Layout ltr = (Layout)acTrans.GetObject((ObjectId)id.Value, OpenMode.ForRead);
                                        layouts.Add(ltr);
                                        editor.WriteMessage(Environment.NewLine + ltr.LayoutName);
                                    }                                                                      
                                }
                                layouts.Sort((l1, l2) => l1.TabOrder.CompareTo(l2.TabOrder));
                                string filename = Path.ChangeExtension(dwgDB.Filename, "pdf");

                                MultiSheetsPdf plotter = new MultiSheetsPdf(filename, layouts);
                                plotter.Publish();
                                editor.WriteMessage(Environment.NewLine + dwgDB.Filename + " успешно экспортирован");
                                acTrans.Commit();
                            }
                        }
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    editor.WriteMessage("\nError: {0}\n{1}", e.Message, e.StackTrace);
                }
            }
        }
    }

}