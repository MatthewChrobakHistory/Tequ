using System.Collections.Generic;
using System.IO;
using Tequ.Data.Models.Maps;

namespace Tequ.Data
{
    public static class DataManager
    {
        public static List<Map> Maps = new List<Map>();

        public static void Load() {
            

            foreach (string file in Directory.GetFiles(Map.MapPath, "*.tmx")) {
                Maps.Add(new Map(file));
            }
        }

        public static void Save() {
            // Include all data-saving logic here.
        }
    }
}
