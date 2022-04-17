using ScriptPortal.Vegas;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCrop
{
    public class EntryPoint
    {
        private Vegas vegas = null;

        public void FromVegas(Vegas vegas)
        {
            this.vegas = vegas;

            var t = FindTrack("Main");
            var zue = t.Events.First(te => te.ActiveTake.Name.Equals("TZoomUP"));
            
        }

        private Track FindTrack(string name)
        {
            return vegas.Project.Tracks.First(t => t.Name == name);
        }
    }
}
