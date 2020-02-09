using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobExporterWF.Models
{
    public class ArborStratix
    {
        // Arbor from Stratix iptfra_rec
        public string Job { get; set; }
        public decimal Wdth {get; set;}
        public int Nbr { get; set; }
    }

    public class ArborExp
    {
        // Arbor expanded from Arbor
        public string Job { get; set; }
        public decimal Wdth { get; set; }
        public int Pos { get; set; }
    }

    public class Consume
    {
        // planned consumtion from iptfrc_rec
        public int Job { get; set; }
        public int Seq { get; set; }
        public string Tag { get; set; }
        public decimal Wdth { get; set; }
        public int Stp { get; set; }
        public int Pos { get; set; }
    }

    public class SO
    {
        // Unique list of SOs on the Job
        public string Trgt { get; set; }
        public string Pfx { get; set; }
        public string Ref { get; set; }
        public string Itm { get; set; }
        public string SItm { get; set; }
    }

    public class Ga
    {
        // Master coil gauge and Job gauge min/max tolerance
        public decimal NumSize { get; set; }
        public decimal GaP { get; set; } = 0;
        public decimal GaN { get; set; } = 0;
    }

    public class HdrFile
    {
        // Detailed information about each SO from List<SO>
        // Stratix transaction common items joined with CPS tolerances
        public string Job { get; set; }
        public string Cust { get; set; } = "Calstrip";
        public string Mtl { get; set; }
        public decimal Wdth { get; set; }
        public decimal Ga { get; set; }
        public decimal KnifeClr { get; set; } = 0.10m;
        public decimal Clr { get; set; }
        public decimal GaP { get; set; }
        public decimal GaN { get; set; }
        public string Note { get; set; } = "None";
    }

    public class MultFile
    {
        // Object that will be exported to EXCEL
        public string Job { get; set; }
        public string Cust { get; set; } = "CalStrip";
        public int Qty { get; set; }
        public decimal Size { get; set; }
        public decimal WdthP { get; set; }
        public decimal WdthN { get; set; }
        public decimal Knife { get; set; } = 0.375m;
    }

    public class MultDetail
    {
        // Object that will be exported to EXCEL
        public string Pfx { get; set; }
        public string Ref { get; set; }
        public string Itm { get; set; }
        public string Cus { get; set; }
        public string Part { get; set; }
        public int CtlNo { get; set; }
        public string Frm { get; set; }
        public decimal Ga { get; set; }
        public decimal GaP { get; set; }
        public decimal GaN { get; set; }
        public decimal Wdth { get; set; }
        public decimal WdthP { get; set; }
        public decimal WdthN { get; set; }
    }

}



