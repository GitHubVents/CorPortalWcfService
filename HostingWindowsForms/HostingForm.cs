using System;
using System.ServiceModel;
using ConecctorOneC;
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
            host = new ServiceHost(typeof(VentsService));
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            host.Open();
            label1.Text = "Ok";
        }
    }
}
