using System;
using System.Windows.Forms;
using System.Management;
using System.Collections;
using System.Text.RegularExpressions;

namespace ComPortIndicator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var timer = new Timer();
            timer.Tick += new EventHandler(this.OnTimer);
            timer.Interval = 500;
            timer.Start();
        }

        private void OnTimer(object sender, EventArgs e)
        {
            string str = GetComPortInfo();
            notifyIcon1.Text = str;
        }

        private string GetComPortInfo()
        {
            var list = new ArrayList();

            ManagementClass win32_pnpentity = new ManagementClass("Win32_PnPEntity");
            ManagementObjectCollection col = win32_pnpentity.GetInstances();

            Regex reg = new Regex(".+\\((?<port>COM\\d+)\\)");

            foreach (ManagementObject obj in col)
            {
                // name : "USB Serial Port(COM??)"
                string name = (string)obj.GetPropertyValue("name");
                if (name != null && name.Contains("(COM"))
                {
                    // "USB Serial Port(COM??)" -> COM??
                    Match m = reg.Match(name);
                    string port = m.Groups["port"].Value;

                    // description : "USB Serial Port"
                    string desc = (string)obj.GetPropertyValue("Description");

                    // result string : "COM?? (USB Serial Port)"
                    list.Add(port);
                }
            }

            if (list.Count == 0)
            {
                return "[No COM port]";
            }

            ComPortComparer comp = new ComPortComparer();
            list.Sort(comp);

            string result = "";
            foreach (string str in list)
            {
                result += str;
                result += "\n";
            }
            return result;
        }

        //
        // see also... http://dobon.net/vb/dotnet/form/hideformwithtrayicon.html
        //
        protected override CreateParams CreateParams
        {
            [System.Security.Permissions.SecurityPermission(
                System.Security.Permissions.SecurityAction.LinkDemand,
                Flags = System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode)]
            get
            {
                const int WS_EX_TOOLWINDOW = 0x80;
                const long WS_POPUP = 0x80000000L;
                const int WS_VISIBLE = 0x10000000;
                const int WS_SYSMENU = 0x80000;
                const int WS_MAXIMIZEBOX = 0x10000;

                CreateParams cp = base.CreateParams;
                cp.ExStyle = WS_EX_TOOLWINDOW;
                cp.Style = unchecked((int)WS_POPUP) |
                    WS_VISIBLE | WS_SYSMENU | WS_MAXIMIZEBOX;
                cp.Width = 0;
                cp.Height = 0;

                return cp;
            }
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            Application.Exit();
        }
    }
}
