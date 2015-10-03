using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestWinForm
{
    public partial class UserControl1 : Form
    {
        public class CtrlmRequest
        {
            public int row;
            public int col;
            public object[,] data;
        }
        public delegate void GetRowColTmpl(ref int row, ref int col, object[,] obj);

        
        private EventHandler sendRequest;
        public GetRowColTmpl getRowCol;
        public UserControl1()
        {
            InitializeComponent();
        }

        public void registerCallback(EventHandler req)
        {
            sendRequest = req;
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            CtrlmRequest r = new CtrlmRequest();
            r.row = Int32.Parse( this.txtRow.Text);
            r.col = Int32.Parse(this.txtCol.Text);
            sendRequest(r, null);
            this.Close();
        }


    }
}
