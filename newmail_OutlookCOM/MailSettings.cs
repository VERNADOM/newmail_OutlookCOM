using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace newmail_OutlookCOM
{
    public partial class MailSettings : Form
    {
        public MailSettings()
        {
            InitializeComponent();
        }

        private void MailSettings_Load(object sender, EventArgs e)
        {
            useNotify();

            emailTextBox.Text = newmail_OutlookCOM.Properties.Settings.Default.Email;
            passwordTextBox.Text = newmail_OutlookCOM.Properties.Settings.Default.Password;

            notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
            notifyIcon1.BalloonTipClicked += new System.EventHandler(this.notifyIcon1_BalloonTipClicked);
        }

        private void useNotify()
        {
            notifyIcon1.Visible = true;
            notifyIcon1.Text = "Mail Notification";

            notifyIcon1.ContextMenuStrip = contextMenuStrip1;
        }

        private void checkMail()
        {
            Imap imap = new Imap("imap-mail.outlook.com", 993, true);
            imap.Authenicate(emailTextBox.Text, passwordTextBox.Text);
            imap.SelectFolder("INBOX");

            notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
            notifyIcon1.BalloonTipTitle = "Mail Notification";
            notifyIcon1.BalloonTipText = "New Mail: " + imap.GetUnseenMessageCount();
            notifyIcon1.ShowBalloonTip(1);
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(emailTextBox.Text))
            {
                MessageBox.Show("Please check your email address!");
            }
            else if (String.IsNullOrEmpty(passwordTextBox.Text))
            {
                MessageBox.Show("Please check your password");
            }
            else
            {
                this.Visible = false;
                checkMail();

                newmail_OutlookCOM.Properties.Settings.Default.Email = emailTextBox.Text;
                newmail_OutlookCOM.Properties.Settings.Default.Password = passwordTextBox.Text;
                newmail_OutlookCOM.Properties.Settings.Default.Save();
            }
        }

        private void notifyIcon1_BalloonTipClicked(object Sender, EventArgs e)
        {
            Process.Start("www.outlook.com");
        }

        private void notifyIcon1_DoubleClick(object Sender, EventArgs e)
        {
            this.Visible = true;
        }

        private void outlookcomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("www.outlook.com");
        }

        private void mailSettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = true;
        }

        private void checkMailToolStripMenuItem_Click(object sender, EventArgs e)
        {
            checkMail();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            checkMail();
        }

        private void outlookCOMLabel_Click(object sender, EventArgs e)
        {
            Process.Start("www.outlook.com");
        }
    }
}
