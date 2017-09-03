using Tequ.Data.Models.Maps.Layers;
using Tequ.Data.Models.Maps.Layers.Tiles;
using Tequ.Data.Models.Maps.TileAttributes;
using TmxSharp;

namespace Tequ.Data.Models.Maps
{
    public class Map
    {
        public static readonly string MapPath = Game.DataPath + "Maps\\";

        public string Name;
        public Layer[] Layers;
        public TileAttribute[] Attributes;
        public int Width { private set; get; }
        public int Height { private set; get; }

        
        public Map(string tmxFile) {
            this.ConvertTMX(TmxMap.Load(tmxFile));
        }

        private void ConvertTMX(TmxMap map) {

            // Set general properties.
            this.Width = map.Width;
            this.Height = map.Height;

            var tilesets = map.Tileset;

            // Load the layers.
            int layerCount = 0;
            for (int i = 0; i < map.Layer.Length; i++) {
                if (!map.Layer[i].Name.StartsWith("attribute")) {
                    layerCount++;
                }
            }
            this.Layers = new Layer[layerCount];
            this.Attributes = new TileAttribute[this.Width * this.Height];


            for (int i = 0; i < map.Layer.Length; i++) {
                this.Layers[i] = new Layer(this.Width, this.Height);
                var layer = this.Layers[i];

                // Get attributes.
                if (map.Layer[i].Name.StartsWith("attribute")) {

                    for (int tileID = 0; tileID < map.Layer[i].Data.Length; tileID++) {
                        int gid = map.Layer[i].Data[tileID].GID;
                        
                        if (this.Attributes[i] == null) {
                            this.Attributes[i] = new TileAttribute();
                        }

                        foreach (var tileset in tilesets) {
                            if (tileset.Name == "mapattributes") {
                                this.Attributes[i].AddAttribute(TileAttribute.GIDToName(gid));
                            } else {
                                gid -= tileset.Image.Width * tileset.Image.Height;
                            }
                        }
                    }

                    continue;
                }
                
                // Get tile information
                for (int tileID = 0; tileID < map.Layer[i].Data.Length; tileID++) {
                    int gid = map.Layer[i].Data[tileID].GID;

                    foreach (var tileset in tilesets) {
                        if (gid - tileset.Image.Width * tileset.Image.Height < 0) {
                            layer.Tiles[tileID] = new Tile(tileset.Name, gid, tileset.Image.Width, tileset.Image.Height);
                            break;
                        } else {
                            gid -= tileset.Image.Width * tileset.Image.Height;
                        }
                    }
                }
            }
        }
    }
}
