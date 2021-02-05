using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace drawAxonsSW
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public int ReturnValue1 { get; set; }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void b1_Click(object sender, EventArgs e)
        {
            this.ReturnValue1 = 1;
            this.Close();
        }

        private void b2_Click(object sender, EventArgs e)
        {
            this.ReturnValue1 = 2;
            this.Close();
        }

        private void b3_Click(object sender, EventArgs e)
        {
            this.ReturnValue1 = 3;
            this.Close();
        }

        private void b4_Click(object sender, EventArgs e)
        {
            this.ReturnValue1 = 4;
            this.Close();
        }

        private void b5_Click(object sender, EventArgs e)
        {
            this.ReturnValue1 = 5;
            this.Close();
        }
    }
}
