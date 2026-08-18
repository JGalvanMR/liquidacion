using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace liquidacion
{
    public partial class DetallePrecio : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);

        public DetallePrecio(string clave, string fecha1, string fecha2)
        {
            InitializeComponent();

            DataTable dtDetalle = new DataTable();

            thisConnection.Open();
            string query = ";WITH TotalUnidadesPorFolio AS ( " +
                "SELECT " +
                    "A.fcn_folio, " +
                    "SUM(A.fcn_num_unidades) AS total_unidades " +
                "FROM " +
                    "tb_det_facturas A " +
                    "INNER JOIN tb_mstr_facturas_nal B ON A.fcn_folio = B.fcn_folio AND A.fcn_tipo = B.fcn_lugar " +
                "where " +
                    "B.fcn_fecha BETWEEN '" + fecha1 + "' AND '" + fecha2 + "' " +
                    "AND B.fcn_estatus <> 'C' " +
                    "AND B.um_clave = 'USD' " +
                    "AND B.fcn_monto <> B.ncr_monto " +
                    "GROUP BY " +
                        "A.fcn_folio " +
            ") " +
            "SELECT " +
                "a.fcn_folio AS Factura, " +
                /*"b.prod_clave AS Clave, " +*/
                "FORMAT(b.fcn_num_unidades, 'N3', 'es-MX') AS Cajas, " +
                "FORMAT(b.fcn_precio_usd, 'C3', 'es-MX') AS Precio, " +
                "FORMAT((b.fcn_num_unidades * b.fcn_precio_usd), 'C3', 'es-MX') AS Importe, " +
                "FORMAT(a.fcn_monto_transporte, 'C3', 'es-MX') AS Flete, " +
                "FORMAT(t.total_unidades, 'N3', 'es-MX') AS CajasTotales, " +
                "FORMAT(ROUND(ISNULL(a.fcn_monto_transporte / NULLIF(t.total_unidades, 0), 0), 2), 'C3', 'es-MX') AS Descontable, " +
                "FORMAT(b.fcn_precio_usd - ROUND(ISNULL(a.fcn_monto_transporte / NULLIF(t.total_unidades, 0), 0), 2), 'C2', 'es-MX') AS PrecioNvo, " +
                "FORMAT(b.fcn_num_unidades * (b.fcn_precio_usd - ROUND(ISNULL(a.fcn_monto_transporte / NULLIF(t.total_unidades, 0), 0), 2)), 'C2', 'es-MX') AS ImporteNvo " +
            "FROM " +
                "tb_mstr_facturas_nal a " +
            "JOIN " +
                "tb_det_facturas b ON a.fcn_folio = b.fcn_folio AND b.fcn_tipo = a.fcn_lugar " +
            "JOIN " +
                "TotalUnidadesPorFolio t ON t.fcn_folio = a.fcn_folio " +
            "WHERE " +
                "a.fcn_fecha BETWEEN '" + fecha1 + "' AND '" + fecha2 + "' " +
                "AND a.fcn_estatus <> 'C' " +
                "AND a.um_clave = 'USD' " +
                "AND b.prod_clave = '" + clave + "' " +
                "AND a.fcn_monto <> a.ncr_monto " +
            "ORDER BY " +
                "b.prod_clave;";
            SqlDataAdapter adapter = new SqlDataAdapter(query, thisConnection);
            adapter.Fill(dtDetalle);

            decimal tot_cajas = 0, tot_importe = 0, tot_nvo_importe = 0;
            foreach (DataRow gr in dtDetalle.Rows)
            {
                tot_cajas = tot_cajas + Convert.ToDecimal(gr["Cajas"].ToString());
                tot_importe = tot_importe + Convert.ToDecimal(gr["Importe"].ToString().Replace("$", ""));
                tot_nvo_importe = tot_nvo_importe + Convert.ToDecimal(gr["ImporteNvo"].ToString().Replace("$", ""));
            }

            dtDetalle.Rows.Add(null, tot_cajas.ToString("###,##0.000"), Math.Round((tot_importe / tot_cajas), 2).ToString("$###,##0.000"), tot_importe.ToString("$###,##0.000"), null, null, null,
                Math.Round((tot_nvo_importe / tot_cajas), 2).ToString("$###,##0.000"), tot_nvo_importe.ToString("$###,##0.000"));

            lblPrecioAnt.Text = Math.Round((tot_importe / tot_cajas), 2).ToString("$###,##0.000");
            lblPrecioAct.Text = Math.Round((tot_nvo_importe / tot_cajas), 2).ToString("$###,##0.000");


            dtgDetalle.DataSource = dtDetalle;

            dtgDetalle.Columns["Cajas"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["Precio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["Importe"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["Flete"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["CajasTotales"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["Descontable"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["PrecioNvo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dtgDetalle.Columns["ImporteNvo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

        }

        private void lblTitulo_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblTitulo_Click(object sender, EventArgs e)
        {

        }
    }
}
