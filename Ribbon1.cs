using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NotesLinkAddIn_x64
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn logic = ThisAddIn.Instance();
            if (logic != null)
            {
                logic.onButtonNotesLink();
            }
        }
    }
}
