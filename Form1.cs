using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Temp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SWDrawer drawer = new SWDrawer();
            drawer.init();

            drawer.newProject(DocumentType.PART);

            drawer.fastCube();

            Body2[] bodies = drawer.getAllBodies();

            foreach (Body2 body in bodies) 
            { 
                drawer.viewBodyFaces(body);
            
            }


        }
    }
}
