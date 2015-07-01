using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Windows.Forms;

namespace HostingWindowsForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private ServiceHost host;
        private void Form1_Load(object sender, EventArgs e)
        {
            host = new ServiceHost(typeof(CorPortalWcfService.Service1));
            host.Open();
        }
    }
}
