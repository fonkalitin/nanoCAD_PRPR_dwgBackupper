
 using App = HostMgd.ApplicationServices;
 using Db = Teigha.DatabaseServices;
 using Ed = HostMgd.EditorInput;
 using Rtm = Teigha.Runtime;
 using Proc = System.Diagnostics.Process; // для запуска explorer чтобы открыть папку

[assembly: Rtm.CommandClass(typeof(Tools.CadCommand))]

namespace Tools
{
        /// <summary> 
        /// Комманды
        /// </summary>
    class CadCommand : Rtm.IExtensionApplication
    {

        public void Initialize()    
            {
  
                App.DocumentCollection dm = App.Application.DocumentManager;

                Ed.Editor ed = dm.MdiActiveDocument.Editor;

                string sCom = "PRPR_backdwg - Бэкап dwg-файла";
                ed.WriteMessage(sCom);

        }

        public void Terminate()
        {
        // Пусто
        }

        /// <summary>
        /// Основная команда для вызова из командной строки
        /// </summary>
        [Rtm.CommandMethod("PRPR_backdwg")]

        /// <summary>
        /// Это основной метод для создания резервной копии dwg-файла в подкаталоге папки его размещения
        /// </summary>
        public static void PRPR_backdwg()
        {
            Db.Database db = Db.HostApplicationServices.WorkingDatabase;
            App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
            Ed.Editor ed = doc.Editor;
            
            // ---- Имена и пути оригинального dwg-файла (открытого)
            string dwgName = doc.Name; // метод получения полного пути и имени текущего dwg-файла (db.Filename; // Альтернативный метод)
            string dwgFileDirPath = Path.GetDirectoryName(dwgName); // Путь до каталога dwg файла (без имени файла) 
            string dwgFileName = Path.GetFileName(dwgName); // Только имя самого dwg файла с расширением
            string dwgFileNameNoExt = Path.GetFileNameWithoutExtension(dwgFileName); // Только имя самого dwg файла без расширения
            string fileExt = Path.GetExtension(dwgFileName); // Только расширение файла (dwg)

            string timeNow = DateTime.Now.ToString().Replace(':', '-'); // Замена двоеточий в формате времени на дефисы
            
            // ---- Имена и пути резервного каталога и файлов
            string backupFolderName = $"_Резервные копии [{dwgFileNameNoExt}]"; //Имя каталога для резервных копий
            string backupFolderFullName = $"{dwgFileDirPath}\\{backupFolderName}"; // полный путь к каталогу резервных копий
            string dwgNewFileName = $"{dwgFileNameNoExt}_backup_{timeNow}{fileExt}"; // сборка имени нового файл-бэкап.dwg
            string dwgNewFullFileName = $"{backupFolderFullName}\\{dwgNewFileName}"; // Сборка итоговой полного пути к файлу с новым файл-бэкап.dwg

            // Проверка существования целевого каталога для бэкапов
            bool isDirExist = Directory.Exists(backupFolderFullName);
            if (!isDirExist) { Directory.CreateDirectory(backupFolderFullName); } // Если каталога нет то создать его

            DialogResult res = MessageBox.Show(
                $"Сохранить резервную копию текущего dwg-файла? {Environment.NewLine}{Environment.NewLine}" +
                $"Будет сохранен в папку: {Environment.NewLine}{backupFolderName}{Environment.NewLine}{Environment.NewLine}" +
                $"Имя резервного файла: {Environment.NewLine}{dwgNewFileName}", 
                "Подтверждение резервирования", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (res == DialogResult.OK)
            {
                db.SaveAs(dwgNewFullFileName, Db.DwgVersion.Current, false); // Непосредственно операция Сохранение файла с текущей версией dwg
            }
            if (File.Exists(dwgNewFullFileName))
            {
                DialogResult result = MessageBox.Show($"Резервная копия успешно создана! {Environment.NewLine}Открыть папку с вашим dwg-чертежом прямо сейчас?", "Выполнено!", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                Proc.Start("explorer.exe", "/select, \"" + backupFolderFullName + "\"");
                }
            }
            else {
                MessageBox.Show($"Ошибка создания файла!{Environment.NewLine}" +
                    $"Резервная копия не была создана", 
                    "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
                
        }

                
    }

}


