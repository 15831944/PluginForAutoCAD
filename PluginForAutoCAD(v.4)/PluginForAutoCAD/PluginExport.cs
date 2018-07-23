using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Acad = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System.IO;
using System.Collections.Generic;
using Autodesk.AutoCAD.Publishing;
using Autodesk.AutoCAD.PlottingServices;
using System.Text;

namespace AutocadPlugin
{
    public class MultiSheetsPdf //Класс создание многолистового PDF файла
    {
        private string dwgFile, pdfFile, dsdFile, outputDir;
        private int sheetNum;
        IEnumerable<Layout> layouts;

        private const string LOG = "publish.log";

        public MultiSheetsPdf(string pdfFile, IEnumerable<Layout> layouts)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            dwgFile = db.Filename;
            this.pdfFile = pdfFile;
            outputDir = Path.GetDirectoryName(this.pdfFile);
            dsdFile = Path.ChangeExtension(this.pdfFile, "dsd");
            this.layouts = layouts;
        }

        public void Publish()
        {
            if (TryCreateDSD())
            {
                Publisher publisher = Acad.Publisher;
                PlotProgressDialog plotDlg = new PlotProgressDialog(false, sheetNum, true)
                {
                    IsVisible = false
                };
                publisher.PublishDsd(dsdFile, plotDlg);
                plotDlg.Destroy();
                File.Delete(dsdFile);
            }
        }

        private bool TryCreateDSD()
        {
            using (DsdData dsd = new DsdData())
            using (DsdEntryCollection dsdEntries = CreateDsdEntryCollection(layouts))
            {
                if (dsdEntries == null || dsdEntries.Count <= 0) return false;

                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                sheetNum = dsdEntries.Count;

                dsd.SetDsdEntryCollection(dsdEntries);
                dsd.SetUnrecognizedData("PwdProtectPublishedDWF", "FALSE");
                dsd.SetUnrecognizedData("PromptForPwd", "FALSE");
                dsd.SheetType = SheetType.MultiDwf;
                dsd.NoOfCopies = 1;
                dsd.DestinationName = pdfFile;
                dsd.IsHomogeneous = false;
                dsd.LogFilePath = Path.Combine(outputDir, LOG);

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
                dsdEntry.DwgName = dwgFile;
                dsdEntry.Layout = layout.LayoutName;
                dsdEntry.Title = Path.GetFileNameWithoutExtension(dwgFile) + "-" + layout.LayoutName;
                dsdEntry.Nps = layout.TabOrder.ToString();
                entries.Add(dsdEntry);
            }
            return entries;
        }

        private void PostProcessDSD(DsdData dsd)
        {
            string str, newStr;
            string tmpFile = Path.Combine(outputDir, "temp.dsd");

            dsd.WriteDsd(tmpFile);

            using (StreamReader reader = new StreamReader(tmpFile, Encoding.Default))
            using (StreamWriter writer = new StreamWriter(dsdFile, false, Encoding.Default))
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
                        newStr = "OriginalSheetPath=" + dwgFile;
                    }
                    else if (str.Contains("Type"))
                    {
                        newStr = "Type=6";
                    }
                    else if (str.Contains("OUT"))
                    {
                        newStr = "OUT=" + outputDir;
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
                        newStr = "LogFilePath=" + Path.Combine(outputDir, LOG);
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
    public class Initialization : IExtensionApplication
    {
        
        public void Initialize()
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Инициализация плагина.." + Environment.NewLine);

        }

        public void Terminate()
        {

        }
     
        [CommandMethod("HelpSaveAsPDF")]
        public void Help()//Команда помощи по использованию плагина (Возможна необходимость корректироваки)
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Для запуска пакетного сохранения листов чертежей в формате .pdf" + Environment.NewLine +
                "наберите в командной строке команду в следующем формате:" + Environment.NewLine +
                "(saveaspdf " + '"' +"диск:/папка1/.../папка с чертежами" +'"' + " "+ '"' + "true/false" + '"' + ")");
        }
    }
    public class Functions //Класс основных функций плагина
    {
        [LispFunction("SaveAsPDF")]
        public static void SaveAsPDF(ResultBuffer rbArgs)//Основная функция плагина (Требуется корректировка и тестирование)
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
                    if (rb.TypeCode == (int)LispDataType.Text)
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
                    DirectoryInfo rootDir = new DirectoryInfo(str);
                    List<FileInfo> files = new List<FileInfo>();
                    short bgp = (short)Acad.GetSystemVariable("BACKGROUNDPLOT");
                    Acad.SetSystemVariable("BACKGROUNDPLOT", 0);
                    GetFiles(rootDir, files, subDirs);
                    //Список листов найденных чертежей
                    foreach (FileInfo file in files)
                    {
                        editor.WriteMessage(Environment.NewLine + file);
                        using (Database dwgDB = new Database(false, true))
                        {
                            dwgDB.ReadDwgFile(file.FullName, FileShare.ReadWrite, false, "");
                           
                            
                            //Далее манипуляции с базой данных чертежа
                            using (Transaction acTrans = dwgDB.TransactionManager.StartTransaction())
                            {
                                List<Layout> layouts = new List<Layout>();
                                DBDictionary layoutDict = (DBDictionary)acTrans.GetObject(dwgDB.LayoutDictionaryId, OpenMode.ForRead);
                                editor.WriteMessage("Листы чертежа " + file + ":");
                                foreach (DBDictionaryEntry id in layoutDict)
                                {
                                    if (id.Key != "Model")
                                    {

                                        Layout ltr = (Layout)acTrans.GetObject((ObjectId)id.Value, OpenMode.ForRead);
                                        PlotConfig config = PlotConfigManager.SetCurrentConfig("AutoCAD PDF (High Quality Print).pc3");
                                        PlotSettingsValidator psv = PlotSettingsValidator.Current;
                                        PlotSettings ps = new PlotSettings(ltr.ModelType);
                                        psv.SetPlotConfigurationName(ps, "AutoCAD PDF (High Quality Print).pc3", null);
                                        psv.RefreshLists(ps);
                                        layouts.Add(ltr);
                                        editor.WriteMessage(Environment.NewLine + ltr.LayoutName);
                                        editor.WriteMessage(Environment.NewLine + ltr.PlotPaperSize.ToString());
                                        editor.WriteMessage(Environment.NewLine + ltr.PlotSettingsName);
                                        
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
                    Acad.SetSystemVariable("BACKGROUNDPLOT", bgp);
                    
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    editor.WriteMessage("\nError: {0}\n{1}", e.Message, e.StackTrace);
                }
            }
            else
            {

            }
        }
        public static void GetFiles(DirectoryInfo rootDir, List<FileInfo> files, bool subdirs) // Получение списка файлов
        {
            files.AddRange(rootDir.GetFiles("*.dwg"));
            if (subdirs)
            {
                DirectoryInfo[] dirs = GetDirectories(rootDir);
                foreach (var dir in dirs)
                {
                    try
                    {
                        GetFiles(dir, files, subdirs);
                    }
                    catch
                    {
                        var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
                        editor.WriteMessage("Каталог " + dir + " недоступен." + Environment.NewLine);
                    }
                }
            }
        }

        private static DirectoryInfo[] GetDirectories(DirectoryInfo rootDir) //Получение списка поддиректорий
        {
            DirectoryInfo[] subdirs;
            subdirs = rootDir.GetDirectories();
            return subdirs;
        }
    }
    public class TestPost //Класс тестов
    {
        [CommandMethod("test1")]
        public void Test()// Тестирование получения всех параметров доступных принтеров
        {
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            var Text = new List<string>();
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            System.Collections.Specialized.StringCollection vs = psv.GetPlotDeviceList();
            foreach (var st in vs)
            {
                editor.WriteMessage(st.ToString() + Environment.NewLine);
                Text.Add(st.ToString());
            }
            using (PlotSettings ps = new PlotSettings(false))
            {
                foreach (var st in vs)
                {
                    editor.WriteMessage("Текущий/Selected: " + st.ToString() + Environment.NewLine);
                    Text.Add("Текущий/Selected: " + st.ToString());
                    if (st.ToString() == "AutoCAD PDF (High Quality Print).pc3")
                    {
                        PlotConfig config = PlotConfigManager.SetCurrentConfig(st.ToString());
                        PlotInfo info = new PlotInfo();
                        psv.SetPlotConfigurationName(ps, st.ToString(), null);
                        psv.RefreshLists(ps);
                        System.Collections.Specialized.StringCollection mediaName = psv.GetCanonicalMediaNameList(ps);
                        foreach (var media in mediaName)
                        {
                            editor.WriteMessage("Формат/Media name " + media.ToString() + Environment.NewLine);
                            Text.Add("Формат/Media name " + media.ToString());
                            MediaBounds bounds = config.GetMediaBounds(media.ToString());
                            editor.WriteMessage(Math.Round(bounds.PageSize.X).ToString() + Environment.NewLine);
                            editor.WriteMessage(Math.Round(bounds.PageSize.Y).ToString() + Environment.NewLine);
                            editor.WriteMessage(Math.Round(bounds.LowerLeftPrintableArea.X).ToString() + Environment.NewLine);
                            editor.WriteMessage(Math.Round(bounds.UpperRightPrintableArea.Y).ToString() + Environment.NewLine);
                            editor.WriteMessage(Math.Round(bounds.UpperRightPrintableArea.X).ToString() + Environment.NewLine);
                            editor.WriteMessage(Math.Round(bounds.LowerLeftPrintableArea.Y).ToString() + Environment.NewLine);
                            Text.Add(Math.Round(bounds.PageSize.X).ToString());
                            Text.Add(Math.Round(bounds.PageSize.Y).ToString());
                            Text.Add(Math.Round(bounds.LowerLeftPrintableArea.X).ToString());
                            Text.Add(Math.Round(bounds.UpperRightPrintableArea.Y).ToString());
                            Text.Add(Math.Round(bounds.UpperRightPrintableArea.X).ToString());
                            Text.Add(Math.Round(bounds.LowerLeftPrintableArea.Y).ToString());


                        }
                    }
                }
            }
            System.IO.File.WriteAllLines("C:\\users\\dyn1\\desktop\\info.txt", Text);
        }
        [CommandMethod("test2")]
        public void Test2()//Тестирование получения списка параметров листов чертежей 
        {
            string path = "C:\\Users\\dyn1\\Desktop";
            
            System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(path);
            System.IO.FileInfo[] dirs = rootDir.GetFiles("*.dwg");
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            var Text = new List<string>();
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            foreach (FileInfo dir in dirs)
            {
                editor.WriteMessage(Environment.NewLine + dir);
                Text.Add(dir.ToString());
                using (Database dwgDB = new Database(false, true))
                {
                    dwgDB.ReadDwgFile(dir.FullName, System.IO.FileShare.ReadWrite, false, "");


                    //Далее манипуляции с базой данных чертежа
                    using (Transaction acTrans = dwgDB.TransactionManager.StartTransaction())
                    {
                        List<Layout> layouts = new List<Layout>();
                        DBDictionary layoutDict = (DBDictionary)acTrans.GetObject(dwgDB.LayoutDictionaryId, OpenMode.ForRead);
                        editor.WriteMessage("Листы чертежа " + dir + ":");
                        Text.Add("Листы чертежа " + dir + ":");
                        foreach (DBDictionaryEntry id in layoutDict)
                        {
                            if (id.Key != "Model")
                            {
                                Layout ltr = (Layout)acTrans.GetObject((ObjectId)id.Value, OpenMode.ForRead);
                                layouts.Add(ltr);
                                editor.WriteMessage(Environment.NewLine + ltr.LayoutName);
                                editor.WriteMessage(Environment.NewLine + Math.Round(ltr.PlotPaperSize.X).ToString());
                                editor.WriteMessage(Environment.NewLine + Math.Round(ltr.PlotPaperSize.Y).ToString());
                                editor.WriteMessage(Environment.NewLine + ltr.PlotSettingsName);
                                Text.Add(ltr.LayoutName);
                                Text.Add(ltr.PlotPaperSize.ToString());
                                Text.Add(Math.Round(ltr.PlotPaperSize.X).ToString());
                                Text.Add(Math.Round(ltr.PlotPaperSize.Y).ToString());
                                Text.Add(ltr.PlotSettingsName);

                            }
                        }
                        System.IO.File.WriteAllLines("C:\\users\\dyn1\\desktop\\info2.txt", Text);
                        acTrans.Commit();
                    }
                }
            }
        }
        [CommandMethod ("test3")]
        public void Test3() // Тестирование получения списка файлов
        {
            string path = "C:\\Users\\dyn1";
            string str = path.Replace('/', '\\');
            System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(str);
            List<FileInfo> files = new List<FileInfo>();
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                GetFiles(rootDir,files,true);
                if (files.Count != 0)
                {
                    foreach (var file in files)
                    {
                       
                        editor.WriteMessage("Файл " + file.Name + " расположен в директории " + file.DirectoryName + Environment.NewLine);
                    }
                }
                else
                {
                    editor.WriteMessage("В каталоге нет чертежей." + Environment.NewLine);
                }
            }
            catch
            {
               
                editor.WriteMessage("Каталог " + rootDir + " недоступен." + Environment.NewLine + "Проверьте верно ли указан путь." + Environment.NewLine);
            }

        }

        public void GetFiles(System.IO.DirectoryInfo rootDir, List<FileInfo> files, bool subdirs) // Получение списка файлов
        {
            files.AddRange(rootDir.GetFiles("*.dwg"));
            if (subdirs)
            {
                System.IO.DirectoryInfo[] dirs = GetDirectories(rootDir);
                foreach(var dir in dirs)
                {
                    try
                    {
                        GetFiles(dir, files, subdirs);
                    }
                    catch
                    {
                        var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
                        editor.WriteMessage("Каталог " + dir + " недоступен." + Environment.NewLine);
                    }
                }
            }
        }

        private System.IO.DirectoryInfo[] GetDirectories(DirectoryInfo rootDir) //Получение списка поддиректорий
        {
            System.IO.DirectoryInfo[] subdirs;
            subdirs = rootDir.GetDirectories();
            return subdirs;
        }
    }
}