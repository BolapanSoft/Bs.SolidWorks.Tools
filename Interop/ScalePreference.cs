using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bs.SolidWorks.Tools.Interop {
    internal struct ScalePreference {
        public double ScaleDecimal { get; set; }
        public bool UseParentScale { get; set; }
        public int UseSheetScale { get; set; }
    }
}
