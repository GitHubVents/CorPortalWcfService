using System;
using System.ServiceModel;

using System.Windows.Forms;

namespace HostingWindowsForms
{
    public partial class HostingForm : Form
    {
        public HostingForm()
        {
            InitializeComponent();
        }

        private ServiceHost host;
        private void Form1_Load(object sender, EventArgs e)
        {
            //host = new ServiceHost(typeof(CorPortalWcfService.Service1));
           // host.Open();
        }
    }
}
