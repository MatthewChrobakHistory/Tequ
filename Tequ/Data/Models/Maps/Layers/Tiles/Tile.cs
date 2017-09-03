namespace Tequ.Data.Models.Maps.Layers.Tiles
{
    public class Tile
    {
        public string Surface { private set; get; }
        public int SurfaceID { private set; get; }
        public int X;
        public int Y;

        public Tile(string surface, int gid, int tilesetWidth, int tilesetHeight) {
            this.Surface = surface;
            this.SurfaceID = -1;
        }

        public void SetSurfaceID(int value) {
            this.SurfaceID = value;
            this.Surface = null;
        }
    }
}
