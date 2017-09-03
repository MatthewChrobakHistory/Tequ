using System;
using SFML.Graphics;

namespace Tequ.Graphics.Sfml.Scenes.Objects
{
    public class Button : SceneObject
    {
        public string Caption;
        public Color TextColor = Color.Black;
        public uint FontSize = 12;

        public override void Draw() {
            // Draw the surface if we have one.
            base.Draw();

            // Draw the button's caption.
            base.RenderCaption(this.Caption, this.FontSize, this.TextColor);
        }

        public sealed override string GetObjectType() {
            return "button";
        }
    }
}
