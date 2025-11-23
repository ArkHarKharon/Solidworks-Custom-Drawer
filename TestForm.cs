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
    public partial class TestForm : Form
    {
        SWDrawer drawer = new SWDrawer();
        public TestForm()
        {
            InitializeComponent();
        }

        private void TestForm_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            drawer.init();
            drawer.connectToOpenedPart();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<object> selected = drawer.GetSelectedObjects();

            HolesArrayCutter cutter = new HolesArrayCutter(drawer);


            cutter.cutHoles(int.Parse(textBox1.Text), int.Parse(textBox2.Text), int.Parse(textBox3.Text), double.Parse(textBox5.Text),double.Parse(textBox4.Text));

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
