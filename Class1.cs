
    using App = HostMgd.ApplicationServices;
    using Db = Teigha.DatabaseServices;
    using Ed = HostMgd.EditorInput;
    using Rtm = Teigha.Runtime;
    using System.Windows.Forms;
    using prPr_FileBackupper;
    using System.IO;
using Multicad;
using mcDBs = Multicad.DatabaseServices;
using Multicad.Objects;
using Multicad.DatabaseServices;
using env = System.Environment;

[assembly: Rtm.CommandClass(typeof(Tools.CadCommand))]

namespace Tools
    {
        /// <summary> 
        /// Комманды
        /// </summary>
        class CadCommand : Rtm.IExtensionApplication
        {

        

        #region INIT
        public void Initialize()    
            {
  
                App.DocumentCollection dm = App.Application.DocumentManager;

                Ed.Editor ed = dm.MdiActiveDocument.Editor;

                string sCom = "PRPR_backdwg" + "Бэкап dwg-файла";
                ed.WriteMessage(sCom);

#if DEBUG
                //для отладки список команд
#endif
            }

            public void Terminate()
            {
                // Пусто
            }

        #endregion

        #region Command

        /// <summary>
        /// Основная команда для вызова из командной строки
        /// </summary>
        [Rtm.CommandMethod("PRPR_backdwg")]


        public static void backupDwg_mode() // Данный метод только вызывает основную форму
        {
            MainForm mf = new MainForm();
            mf.ShowDialog();
                    
        }


        /// <summary>
        /// Это основной метод для обработки, он принимает значение номера ревизии (RevisionValue) из формы
        /// </summary>
        public static void backupDwg_mode_1(string RevisionValue)
            {
                      
                Db.Database db = Db.HostApplicationServices.WorkingDatabase;
                App.Document doc = App.Application.DocumentManager.MdiActiveDocument;
                Ed.Editor ed = doc.Editor;


            //string fName = db.Filename; // Альтернативный метод получения полного пути и имени текущего dwg-файла
           
            string dwgName = doc.Name; // метод получения полного пути и имени текущего dwg-файла
            int pos = dwgName.LastIndexOf("\\");
            string dwgPath = dwgName.Remove(pos, dwgName.Length - pos);
            string[] dwgSplitName = dwgName.Split(new char[] { '.' }); // Отделение расширения *.dwg от полного имени файла
            string timeNow = DateTime.Now.ToString().Replace(':', '-'); // Замена двоеточий в формате времени на дефисы
            string dwgNewName = dwgSplitName[0] + "_back_" + RevisionValue + "_" + timeNow + ".dwg"; // Сборка итоговой полной строки с путем к файлу с новым именем
            
            //DialogResult res = MessageBox.Show("Вы действительно хотите сделать резервную копию текущего dwg-файла? " + Environment.NewLine + dwgNewName, "Подтверждение резервирования", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            //if (res == DialogResult.OK){
                db.SaveAs(dwgNewName, Db.DwgVersion.Current, false);

            if (File.Exists(dwgNewName))
            {
                MessageBox.Show("Резервная копия успешно создана!");
            }
            else { 
                MessageBox.Show("Ошибка создания файла!", "Ошибка!"); 
            }
                
}

                
}



        }


            #endregion

