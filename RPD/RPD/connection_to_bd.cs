using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using System.Diagnostics;

namespace RPD
{
    public class connection_to_bd
    {
        public OleDbConnection con;
        public OleDbCommand command = new OleDbCommand();
        public OleDbDataReader reader ;

        
        public void Connect()
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + Application.StartupPath + "/baza_dan_proekt_kh.accdb");
            command.Connection = con;
            con.Open();
        }
       
    }
}
