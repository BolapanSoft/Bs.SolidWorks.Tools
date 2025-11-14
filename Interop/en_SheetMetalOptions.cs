using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bs.SolidWorks.Interop {
    [Flags]
    public enum en_SheetMetalOptions {
        FlatPatternGeometry = 1 << 0,
        HiddenEdges = 1 << 1,
        ExportBendLines = 1 << 2,
        IncludeSketches = 1 << 3,
        MergeCoplanarFaces = 1 << 4,
        ExportLibraryFeatures = 1 << 5,
        ExportFormingTools = 1 << 6,
        ExportBoundingBoxes = 1 << 12,
    }
}
