using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Acad = Autodesk.AutoCAD.ApplicationServices.Core.Application;
using System.IO;
using System.Collections;

namespace AutocadPlugin
{
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
                "(saveaspdf " + '"' +"диск:/папка1/.../папка с чертежами" +'"' + ")");
        }
    }
    public class test1
    {
        [LispFunction("SaveAsPDF")]
        public static void SaveAsPDF(ResultBuffer rbArgs)
        {
            var editor = Acad.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage("Процесс экспорта чертежей в формат PDF начинается...");
            if (rbArgs != null)
            {
                string str = "";
                foreach (TypedValue rb in rbArgs)
                {
                    if (rb.TypeCode == (int)Autodesk.AutoCAD.Runtime.LispDataType.Text)
                        str += rb.Value.ToString();
                }
                Acad.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Путь к чертежам: " + str);
                try
                {
                    str = str.Replace('/', '\\');
                    System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(str);
                    System.IO.FileInfo[] dirs = rootDir.GetFiles("*.dwg");

                    foreach (FileInfo dir in dirs)
                    {
                        editor.WriteMessage(Environment.NewLine + dir);
                        using (Database dwgDB = new Database(false, true))
                        {
                            dwgDB.ReadDwgFile(dir.FullName, System.IO.FileShare.ReadWrite, false, "");
                            //Далее манипуляции с базой данных чертежа
                            using (Transaction acTrans = dwgDB.TransactionManager.StartTransaction())
                            {
                                DBDictionary layoutDict = (DBDictionary)acTrans.GetObject(dwgDB.LayoutDictionaryId, OpenMode.ForRead);
                                editor.WriteMessage("Листы чертежа " + dir + ":");
                                foreach (DictionaryEntry id in layoutDict)
                                {
                                    Layout ltr = (Layout)acTrans.GetObject((ObjectId)id.Value, OpenMode.ForRead);
                                    editor.WriteMessage(Environment.NewLine + ltr.LayoutName);
                                }
                            }
                        }
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    editor.WriteMessage("process failed");
                }
            }
        }
    }

}