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


            drawer.selectDefaultPlane(DefaultPlaneName.TOP);
            drawer.insertSketch(true);
            drawer.createCenterRectangle(0, 0, 0, 1, 0.5, 0);

            drawer.extrude(0.1);


            Body2[] bodies = drawer.getAllBodies();
            Face2[] faces = drawer.getAllFaces(bodies[0]);

            drawer.SelectFaceByIndex(bodies[0],4);
            drawer.insertSketch(true);

            drawer.createCircleByRadius(0,0,0.1,0.3);

            drawer.extrude(0.2);

            bodies = drawer.getAllBodies(); //обновляю bodies, тк добавилось новое тело

            drawer.SelectFaceByIndex(bodies[0],7);
            drawer.insertSketch(true);

            drawer.createCircleByRadius(0, 0, 0.4, 0.2);

            drawer.cutHole(HoleType.DISTANCE,false,0.2);



        }
    }
}
