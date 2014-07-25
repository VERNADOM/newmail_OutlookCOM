using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;
using System.Net.Security;
using System.IO;

namespace newmail_OutlookCOM
{

    class Imap
    {
        protected TcpClient _tcpClient;
        protected StreamReader _reader;
        protected StreamWriter _writer;

        protected string _selectedFolder = string.Empty;
        protected int _prefix = 1;

        public Imap(string host, int port, bool ssl = false)
        {
            try
            {
                _tcpClient = new TcpClient(host, port);

                if (ssl)
                {
                    var stream = new SslStream(_tcpClient.GetStream());
                    stream.AuthenticateAsClient(host);

                    _reader = new StreamReader(stream);
                    _writer = new StreamWriter(stream);
                }
                else
                {
                    var stream = _tcpClient.GetStream();
                    _reader = new StreamReader(stream);
                    _writer = new StreamWriter(stream);
                }

                string greeting = _reader.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void Authenicate(string user, string pass)
        {
            this.SendCommand(string.Format("LOGIN {0} {1}", user, pass));
            string response = this.GetResponse();
        }

        public void SelectFolder(string folderName)
        {
            this._selectedFolder = folderName;
            this.SendCommand("SELECT " + folderName);
            string response = this.GetResponse();
        }

        public int GetUnseenMessageCount()
        {
            this.SendCommand("STATUS " + this._selectedFolder + " (unseen)");
            string response = this.GetResponse();
            Match m = Regex.Match(response, "[0-9]*[0-9]");
            return Convert.ToInt32(m.ToString());
        }

        protected void SendCommand(string cmd)
        {
            _writer.WriteLine("A" + _prefix.ToString() + " " + cmd);
            _writer.Flush();
            _prefix++;
        }

        protected string GetResponse()
        {
            string response = string.Empty;

            while (true)
            {
                string line = _reader.ReadLine();
                string[] tags = line.Split(new char[] { ' ' });
                response += line + Environment.NewLine;
                if (tags[0].Substring(0, 1) == "A" && tags[1].Trim() == "OK" || tags[1].Trim() == "BAD" || tags[1].Trim() == "NO")
                {
                    break;
                }

            }

            return response;
        }
    }

}