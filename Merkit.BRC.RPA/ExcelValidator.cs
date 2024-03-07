using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Merkit.RPA.PA.Framework;

namespace Merkit.BRC.RPA
{

    /// <summary>
    /// BRC_Enterhungary input excel ellenőrzése
    /// </summary>
    public class ExcelValidator
    {
        public string TextFilePath { get; set; }

        public ExcelValidator()
        {

        }

        public ExcelValidator(string textFilePath)
        {
            this.TextFilePath = textFilePath;
        }

        string állampolgárság_dropdown = "";
        string átvételi_ország_dropdown = "";
        string benyújtó_dropdown = "";
        string családi_állapot_dropdown = "";
        string egészségbiztosítás_dropdown = "";

        /// <summary>
        /// Az oldalon lévő , a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett  txt fájlokból változókba
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool LoadDropdownValues(string path)
        {

            // 22 ?
            állampolgárság_dropdown = FileManager.ReadTextFile(Path.Combine(path, "állampolgárság.txt"));
            átvételi_ország_dropdown = FileManager.ReadTextFile(Path.Combine(path, "átvételi ország.txt"));
            benyújtó_dropdown = FileManager.ReadTextFile(Path.Combine(path, "benyújtó.txt"));
            családi_állapot_dropdown = FileManager.ReadTextFile(Path.Combine(path, "családi állapot.txt"));
            egészségbiztosítás_dropdown = FileManager.ReadTextFile(Path.Combine(path, "egészségbiztosítás.txt"));

            return true;
        }    

    }
}
