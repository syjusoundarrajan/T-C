using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wall_Rect
{
    public class Class1
    {
        public Point EndLocation
        {
            get
            {
                int x;
                int y;

                switch (this.Orientation)
                {
                    case SegmentOrientation.Vertical:
                        x = this.Location.X;
                        y = this.Location.Y + this.Size;
                        break;
                    default:
                        x = this.Location.X + this.Size;
                        y = this.Location.Y;
                        break;
                }

                return new Point(x, y);
            }
        }

        public Point Location { get; set; }

        public SegmentOrientation Orientation { get; set; }

        public int Size { get; set; }
    }

    internal class SegmentPoint
    {
        public SegmentPointConnections Connections { get; set; }

        public Point Location { get; set; }

        public int X { get { return this.Location.X; } }

        public int Y { get { return this.Location.Y; } }
    }
}
