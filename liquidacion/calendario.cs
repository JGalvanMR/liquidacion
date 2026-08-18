using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace liquidacion
{
    public partial class calendario : Form
    {
        public calendario()
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);
        }

        public class datosval
        {
            private string _fecha;
            public string fecha
            {
                get { return _fecha; }
                set { _fecha = value; }
            }
        }
        public class SharedData
        {
            public static datosval Polino;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            //MessageBox.Show(monthCalendar1.SelectionRange.Start.Date.ToShortDateString());
            datosval passdata = new datosval();
            passdata.fecha = monthCalendar1.SelectionRange.Start.Date.ToShortDateString();
            SharedData.Polino = passdata;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void calendario_Load(object sender, EventArgs e)
        {

        }
    }
}
