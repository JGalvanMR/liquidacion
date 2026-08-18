using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace liquidacion
{
    public partial class anticipos : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1;
        SqlCommand cmnd1;

        DataTable dt = new DataTable();

        public anticipos(string anti, string cve_prov, string nom_prov)
        {
            InitializeComponent();

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            string ant = anti;

            dt.Columns.Add("liq_folio", typeof(string));
            dt.Columns.Add("liq_tipo", typeof(string));
            dt.Columns.Add("tipo_mov", typeof(string));
            dt.Columns.Add("descuento", typeof(string));

            lblProveedor.Text = nom_prov;

            thisConnection.Open();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT A.Fecha, A.Id_Contrato, A.Descripcion_Art, A.Cantidad, A.contrato, A.factura, B.prod_nombre FROM Tb_Prestamos_Prov A JOIN " +
                "tb_cat_producto B ON B.prod_clave = A.Id_Contrato WHERE Id_Movimiento = '" + anti + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    lblFecha.Text = Convert.ToDateTime(reader1["Fecha"].ToString()).ToShortDateString();
                    lblCultivo.Text = reader1["prod_nombre"].ToString().Trim();
                    lblAnticipo.Text = reader1["Descripcion_Art"].ToString().Trim();
                    lblCantidad.Text = Convert.ToDecimal(reader1["Cantidad"].ToString()).ToString("$###,###,##0.0000");
                    lblContrato.Text = reader1["contrato"].ToString().Trim();
                    lblFactura.Text = reader1["factura"].ToString().Trim();
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT liq_folio, cantidad, liq_tipo, tipo_mov, tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + anti + "' AND estatus = 'A'";
            reader1 = cmnd1.ExecuteReader();
            decimal cantidad = 0;
            DataRow rt;
            decimal sumatoria = 0;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rt = dt.NewRow();
                    rt["liq_folio"] = reader1["liq_folio"].ToString().Trim();
                    rt["liq_tipo"] = reader1["liq_tipo"].ToString().Trim();
                    rt["tipo_mov"] = reader1["tipo_mov"].ToString().Trim();

                    decimal var1 = Convert.ToDecimal(reader1["tipo_cambio"]);
                    string var2 = reader1["liq_tipo"].ToString().Trim();
                    if (var2 == "N" && var1 == 1)//anticipo en pesos
                    {
                        cantidad = cantidad + Convert.ToDecimal(reader1["cantidad"]);
                        sumatoria = sumatoria + cantidad;
                    }
                    if (var2 == "N" && var1 > 1)//anticipo en dolares
                    {
                        decimal dato1 = Convert.ToDecimal(reader1["cantidad"]) / Convert.ToDecimal(reader1["tipo_cambio"]);
                        cantidad = cantidad + Math.Round(dato1, 4);
                        sumatoria = sumatoria + cantidad;
                    }
                    if (var2 == "E" && var1 == 1)//anticipo en dolares
                    {
                        cantidad = cantidad + Convert.ToDecimal(reader1["cantidad"]);
                        sumatoria = sumatoria + cantidad;
                    }
                    if (var2 == "E" && var1 > 1)//anticipo en pesos
                    {
                        decimal dato1 = Convert.ToDecimal(reader1["cantidad"]) * Convert.ToDecimal(reader1["tipo_cambio"]);
                        cantidad = cantidad + Math.Round(dato1, 4);
                        sumatoria = sumatoria + cantidad;
                    }

                    rt["descuento"] = cantidad.ToString();
                    dt.Rows.Add(rt);
                    cantidad = 0;
                }
                rt = dt.NewRow();
                rt["liq_folio"] = "";
                rt["liq_tipo"] = "";
                rt["tipo_mov"] = "TOTAL";
                rt["descuento"] = sumatoria.ToString();
                dt.Rows.Add(rt);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            lblDescuento.Text = Convert.ToDecimal(sumatoria).ToString("$###,###,##0.0000");

            string sald = lblCantidad.Text.Replace("$", "");
            decimal saldo = Convert.ToDecimal(sald) - sumatoria;
            lblSaldo.Text = Convert.ToDecimal(saldo).ToString("$###,###,##0.0000");

            foreach (DataRow r in dt.Rows)
            {
                string liq_tipo = (r["liq_tipo"].ToString() == "N") ? "NACIONAL" : (r["liq_tipo"].ToString() == "N") ? "EXPORTACION" : "";
                string tipo_mov = (r["tipo_mov"].ToString() == "LQ") ? "LIQUIDACION" : (r["tipo_mov"].ToString() == "MP") ? "OC MAT PRIMA" : (r["tipo_mov"].ToString() == "ES") ? "LIQ. ESPARRAGO" : "TOTAL";
                dtgMovimientos.Rows.Add(r["liq_folio"].ToString(), liq_tipo, tipo_mov, Convert.ToDecimal(r["descuento"]).ToString("$###,###,##0.0000"));
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void anticipos_Load(object sender, EventArgs e)
        {

        }
    }
}
