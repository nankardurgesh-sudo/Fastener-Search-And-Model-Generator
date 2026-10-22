
using NXOpen;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Effects;

namespace SmartFastener
{
    public class Execute
    {
        public static void Main(string[] args)
        {
            SmartFastenerUI smartFastenerUI = new SmartFastenerUI();
            smartFastenerUI.Show();
        }
        public static int GetUnloadOption(string dummy)
        {
            return (int)NXOpen.Session.LibraryUnloadOption.AtTermination;
        }
    }
}
