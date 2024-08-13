using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenTK;
using OpenTK.Graphics;
using OpenTK.Graphics.OpenGL;
using OpenTK.Input;


namespace Yer_İstasyonu_Uygulaması
{
    internal class rocket : GameWindow
    {
        public rocket()
        : base(800, 600, GraphicsMode.Default, "Rocket Simulation")
        {
        }
    }
}
