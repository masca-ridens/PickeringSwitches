using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Ivi.Visa.Interop;
using System.Runtime.InteropServices;

namespace PickeringSwitches
{
    public partial class Form1 : Form
    {
        private readonly ResourceManager ioMgr = new ResourceManager();
        List<FormattedIO488> instruments = new List<FormattedIO488>();
        List<string> instrumentLog = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            if (!instrumentLog.Contains(b.Tag.ToString()))
            {
                instruments.Add(new FormattedIO488());
                instrumentLog.Add(b.Tag.ToString());

                GroupBox gb = b.Parent as GroupBox;
                var tb = gb.Controls.OfType<TextBox>().FirstOrDefault(r => r.Tag.ToString() == "SCPI");
                var label = gb.Controls.OfType<Label>().FirstOrDefault(r => r.Tag != null && r.Tag.ToString() == "idn?");
                string address = tb.Text;

                instruments.Add(new FormattedIO488());
                instrumentLog.Add(b.Tag.ToString());

                try
                {
                    IVisaSession session;// = null;
                    session = ioMgr.Open(address, AccessMode.NO_LOCK, 3000, "");

                    instruments[instruments.Count - 1].IO = (IMessage)session;
                    instruments[instruments.Count - 1].IO.SendEndEnabled = false;
                    instruments[instruments.Count - 1].IO.Timeout = 9000;
                    instruments[instruments.Count - 1].IO.TerminationCharacterEnabled = true;
                    instruments[instruments.Count - 1].WriteString("*IDN?");
                    label.Text = instruments[instruments.Count - 1].ReadString();
                    b.Enabled = false;
                }
                catch (COMException ex)
                {
                    MessageBox.Show("Failed to connect to the SCPI interface." + Environment.NewLine
                    + "Details:" + Environment.NewLine
                    + ex.Message);
                    instruments.Clear();
                    instrumentLog.Clear();
                    return;
                }
            }
            string card = numericUpDown1.Value.ToString();
            string state = radioButton1.Checked ? "Open" : "Close";
            string switchNo = numericUpDown2.Value.ToString();
            instruments[instruments.Count - 1].WriteString(state + " " + card + "," + switchNo);

            b.Enabled = true;
        }
    }
}
