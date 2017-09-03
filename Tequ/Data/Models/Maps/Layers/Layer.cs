using Tequ.Data.Models.Maps.Layers.Tiles;

namespace Tequ.Data.Models.Maps.Layers
{
    public class Layer
    {
        public Tile[] Tiles;
        private int _width;
        private int _height;

        public Layer(int width, int height) {
            this.Tiles = new Tile[width * height];
        }

        public Tile GetTile(int x, int y) {
            return Tiles[y * _width + x];
        }
    }
}
