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
    public partial class cajas : Form
    {
        public cajas()
        {
            InitializeComponent();

            txtCantidad.Focus();
        }

        public class datosval
        {
            private string _cantidad;
            public string cantidad
            {
                get { return _cantidad; }
                set { _cantidad = value; }
            }
        }

        public class SharedData
        {
            public static datosval Polino;
        }

        private void btnCierra_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void btnCantidad_Click(object sender, EventArgs e)
        {
            if (validanumero(txtCantidad.Text) == true)
            {
                datosval passdata = new datosval();
                passdata.cantidad = txtCantidad.Text;
                SharedData.Polino = passdata;
                this.DialogResult = DialogResult.Yes;
                this.Close();
            }
            else
            {
                MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        private void txtCantidad_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (validanumero(txtCantidad.Text) == true)
                {
                    datosval passdata = new datosval();
                    passdata.cantidad = txtCantidad.Text;
                    SharedData.Polino = passdata;
                    this.DialogResult = DialogResult.Yes;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
        }

        private bool validanumero(string value)
        {
            try
            {
                decimal val = Convert.ToDecimal(value);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void cajas_Load(object sender, EventArgs e)
        {

        }
    }
}
