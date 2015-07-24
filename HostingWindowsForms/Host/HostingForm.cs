using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.ServiceModel;
using ConecctorOneC;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;
using System.ServiceModel.Channels;
using System.IO;
using System.Linq;
using System.Xml;
using EdmLib;
using HostingWindowsForms.Data;

namespace HostingWindowsForms
{
    public partial class HostingForm : Form
    {
        RegistryKey rkApp = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

        //Initialize
        public HostingForm()
        {
            InitializeComponent();

            rkApp.SetValue("Vents service", Application.ExecutablePath.ToString());

            try
            {
                var perm = new SqlClientPermission(System.Security.Permissions.PermissionState.Unrestricted);
                perm.Demand();
            }
            catch
            {
                throw new ApplicationException("No permission");
            }

            CheckPdmVault();

        }

        #region Fields
        
        private ServiceHost host;
        private ClassOfTasks _classOdTasks;
        public EPDM Epdm;

        static public IEdmVault5 Vault1;
        static public IEdmVault7 Vault2;

        //public static string VaultName = @"Tets_debag";
        public static string VaultName = @"Vents-PDM";

        #endregion

        private void HostingForm_Load(object sender, EventArgs e)
        {
            try
            {
                host = new ServiceHost(typeof(VentsService));
                host.Open();

                Connection.ConnectionString();

                _classOdTasks = new ClassOfTasks();
                _classOdTasks.OnNewMessageChange += new ClassOfTasks.NewMessage(OnNewMessage);

                //LoadMessages();
                LoadMessagesList();

                labelRun.Text = @"Работает...";

                toolStripMenuRunService.Enabled = false;
             
                //myTimer.Tick += new EventHandler(TimerEventProcessor);

                ////myTimer.Interval = 5000;
                //myTimer.Start();

                //// Runs the timer
                //while (exitFlag == false)
                //{
                //    Application.DoEvents();
                //}
                //myTimer.Stop();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #region Timer

        static Timer myTimer = new Timer();
        static bool exitFlag = false;
        // This is the method to run when the timer is raised.
        private static void TimerEventProcessor(Object myObject, EventArgs myEventArgs)
        {

            (new System.Threading.Thread(delegate ()
            {
                var hour = 8;
                var minute = 04;

                m:
                while (true)
                {
                    // Displays a message box asking whether to continue running the timer.
                    //if (MessageBox.Show("Continue running?", "Count is: " + alarmCounter, MessageBoxButtons.YesNo) == DialogResult.Yes)
                    if ((hour == System.DateTime.Now.Hour) && (minute == System.DateTime.Now.Minute))
                    {
                        MessageBox.Show("1");
                    // Restarts the timer and increments the counter.
                    //alarmCounter += 1;
                    //myTimer.Enabled = true;

                    //return;
                    goto m;
                    }
                    else
                    {
                        // Stops the timer.
                        exitFlag = true;
                    }

                }

            })).Start();       
        }

        private void SetTimer()
        {
            //timer1.Stop();
            //var timeToAlarm = DateTime.Now.Date.AddHours(12).AddMinutes(03);
            //if (timeToAlarm == DateTime.Now)
            //{
            //    MessageBox.Show("1");
            //}
            //    timeToAlarm.AddDays(1);
            //timer1.Interval = (int)(timeToAlarm - DateTime.Now).TotalMilliseconds;
            //timer1.Start();

        }
        #endregion

        #region Tray icon menu

        private void toolStripMenuRunService_Click(object sender, EventArgs e)
        {

            host = new ServiceHost(typeof(VentsService));
            host.Open();
            toolStripMenuRunService.Enabled = false;
            labelRun.Text = "Работает...";

        }

        private void toolStripMenuStopService_Click(object sender, EventArgs e)
        {
            host.Close();
            toolStripMenuRunService.Enabled = true;
            labelRun.Text = "Служба остановлена";
        }

        private void mynotifyicon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();

            this.WindowState = FormWindowState.Normal;

            this.ShowInTaskbar = true;
        }

        private void ToolStripMenuItemExit_Click(object sender, EventArgs e)
        {
            host.Close();
            this.Close();
        }

        #endregion

        public void OnNewMessage()
        {
            var i = (ISynchronizeInvoke)this;

            // Check if the event was generated from another
            // thread and needs invoke instead
            if (i.InvokeRequired)
            {
                var tempDelegate = new ClassOfTasks.NewMessage(OnNewMessage);
                i.BeginInvoke(tempDelegate, null);
                return;
            }

            // If not coming from a seperate thread
            // we can access the Windows form controls

            //LoadMessages();
            LoadMessagesList();
        }
        public List<string> LoadMessagesList()
        {
            _classOdTasks = new ClassOfTasks();

            var listStr = new List<string>();
            listStr.Clear();

            var dt = _classOdTasks.GetTaskListSql();

            listStr.AddRange(from DataRow row in dt.Rows select $"{row["FolderPath"]} : {row["FileName"]}");

            return listStr;
        }

        #region EPDM
        //private const string VaultName = @"Vents-PDM";
        public void CheckPdmVault()
        {
            if (Vault1 == null)
            {
                Vault1 = new EdmVault5();
            }

            Vault2 = (IEdmVault7)Vault1;

            var ok = Vault1.IsLoggedIn;

            if (!ok)
            {
                Vault1.LoginAuto(VaultName, 0);
            }
        }
        #endregion

        #region Buttons

        private void BtnStart_Click(object sender, EventArgs e)
        {
            try
            {
                host.Close();

                host = new ServiceHost(typeof(VentsService));

                host.Open();

                
                _classOdTasks = new ClassOfTasks();
                _classOdTasks.OnNewMessageChange += new ClassOfTasks.NewMessage(OnNewMessage);

                //LoadMessages();

                LoadMessagesList();

                labelRun.Text = @"Работает...";

                toolStripMenuRunService.Enabled = false;
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message); }
        }

        private void BtnStopService_Click(object sender, EventArgs e)
        {
            host.Close();
            toolStripMenuRunService.Enabled = true;
            labelRun.Text = "Служба остановлена";
        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var classTask = new ClassOfTasks { HostForm = this };
            //_classOdTasks = new ClassOfTasks();
            classTask.RunTask();
        }
    }

}
