using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ScriptPortal.Vegas;

namespace MiniScore
{
    public class EntryPoint
    {

        private Vegas vegas = null;

        public void FromVegas(Vegas vegas)
        {
            this.vegas = vegas;

        }

        private string SelectFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                var res = ofd.ShowDialog();
                if (res == DialogResult.OK)
                {
                    return ofd.FileName;
                }
            }
            return String.Empty;
        }

    }
}

