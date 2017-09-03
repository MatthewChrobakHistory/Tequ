using System.Collections.Generic;

namespace Tequ.Data.Models.Maps.TileAttributes
{
    public class TileAttribute
    {
        private List<string> _attributes = new List<string>();

        public void AddAttribute(string attribute) {
            if (!HasAttribute(attribute)) {
                this._attributes.Add(attribute);
            }
        }

        public bool HasAttribute(string attribute) {
            return this._attributes.Contains(attribute);
        }

        public static string GIDToName(int gid) {
            switch (gid) {
                default:
                    return string.Empty;
            }
        }
    }
}
