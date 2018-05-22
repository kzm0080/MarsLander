using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarsLander
{
   public class Discrete
    {
        public int Iteration { get; set; }
        public double X { get; set; }
        public double SumX { get; set; }
        public double Y { get; set; }
        public double SumY { get; set; }
        public double FallProbability { get; set; }
        public string Continue { get; set; }
        public string Angle { get; set; }
        public double Distance { get; set; }
    }
}
