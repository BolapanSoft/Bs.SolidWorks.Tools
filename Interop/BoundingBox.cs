using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;

namespace Bs.SolidWorks.Tools.Interop {
    internal struct BoundingBox : IEquatable<BoundingBox> {
        public double Xmin { get; } = 0;
        public double Ymin { get; } = 0;
        public double Xmax { get; } = 0;
        public double Ymax { get; } = 0;

        public BoundingBox(double xMin, double yMin, double xMax, double yMax) {
            Xmin = xMin;
            Ymin = yMin;
            Xmax = xMax;
            Ymax = yMax;
        }

        public BoundingBox(double[]? obj) {
            if (obj is not null) {
                if (obj.Length != 4)
                    throw new ArgumentOutOfRangeException(nameof(obj));
                int i = 0;
                Xmin = obj[i++];
                Ymin = obj[i++];
                Xmax = obj[i++];
                Ymax = obj[i++];
            }
        }

        /// <summary>
        /// Deconstruct to enable tuple-like unpacking: var (xmin, ymin, xmax, ymax) = box;
        /// </summary>
        public void Deconstruct(out double xmin, out double ymin, out double xmax, out double ymax) {
            xmin = Xmin;
            ymin = Ymin;
            xmax = Xmax;
            ymax = Ymax;
        }

        public override bool Equals(object obj) => obj is BoundingBox other && Equals(other);

        public bool Equals(BoundingBox other) {
            return Xmin == other.Xmin &&
                   Ymin == other.Ymin &&
                   Xmax == other.Xmax &&
                   Ymax == other.Ymax;
        }

        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = hash * 23 + Xmin.GetHashCode();
                hash = hash * 23 + Ymin.GetHashCode();
                hash = hash * 23 + Xmax.GetHashCode();
                hash = hash * 23 + Ymax.GetHashCode();
                return hash;
            }
        }

        public static bool operator ==(BoundingBox left, BoundingBox right) => left.Equals(right);
        public static bool operator !=(BoundingBox left, BoundingBox right) => !(left == right);

        public override string ToString() => $"({Xmin}, {Ymin}, {Xmax}, {Ymax})";
    }

}
