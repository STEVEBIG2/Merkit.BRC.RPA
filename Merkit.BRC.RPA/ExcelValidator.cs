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
    public static class ExcelValidator
    {
        public static string TextFilePath { get; set; }

        public static Dictionary<string, string> loadDropdownDict = new Dictionary<string, string>();

        /// <summary>
        /// Az oldalon lévő , a flowhoz szükséges dropdown elemek értékeit  betölti a  lementett txt fájlokból
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool LoadDropdownValues(string path)
        {
            loadDropdownDict.Clear();

            string[] dropdownType = {
                "állampolgárság", "átvételi ország", "benyújtó", "családi állapot", "egészségbiztosítás",
                "előző ország", "FEOR", "iskolai végzettség", "munkakör iskolai végzettség", "munkáltató közterület jellege",
                "nem", "nemzetiség", "nyelv", "pénznem", "szállás emelet",
                "szállás közterület jellege", "szállás tartózkodási jogcíme", "szül_ország", "TEÁOR", "továbbut ország",
                "útlevél tipus", "zipcode"
            };

            foreach (string type in dropdownType)
            {
                loadDropdownDict.Add(
                    String.Format("{0}_dropdown", type.Replace(" ", "_")),
                    FileManager.ReadTextFile(Path.Combine(path, type + ".txt"))
                    );
            }

            return true;
        }    

    }
}
