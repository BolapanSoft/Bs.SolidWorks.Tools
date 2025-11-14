using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bs.SolidWorks.Tools {
    internal class Resources {
        public string GostDrwTemplate = @"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\gost-part drw.drwdot";
        public static Resources Instance = new Resources();
    }
}
