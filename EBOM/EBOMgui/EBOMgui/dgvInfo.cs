using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;


namespace EBOMgui
{
    class dgvInfo
    {
        public List<int> rows, columns;
        public List<string> text;
        public List<Color> colors;

        public dgvInfo()
        {
            rows = new List<int>();
            columns = new List<int>();
            text = new List<string>();
            colors = new List<Color>();
        }
    }
}
