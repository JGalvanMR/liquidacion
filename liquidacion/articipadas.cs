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
    public partial class articipadas : Form
    {
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1;
        SqlCommand cmnd1;

        string cv_prov = "";
        string proveedor = "";
        string cv_prod = "";
        string producto = "";
        string cantidad = "";
        string tipo = "";

        string fecha1 = "";
        string fecha2 = "";

        public articipadas(string cveprov, string prov, string cveprod, string prod, string cant, string tip, string f1, string f2)
        {
            InitializeComponent();

            cv_prov = cveprov;
            proveedor = prov;
            cv_prod = cveprod;
            producto = prod;
            cantidad = cant;
            tipo = tip;

            fecha1 = f1;
            fecha2 = f2;

            lblCveProv.Text = cv_prov;
            lblProveedor.Text = proveedor;
            lblClave.Text = cv_prod;
            lblNombre.Text = producto;
            lblCantidad.Text = cantidad;

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();

            cmnd1.CommandText = "SELECT A.numero_oc, FORMAT(A.fecha_oc, 'dd-MM-yyyy') AS fecha_oc, RTRIM(B.prod_clave) AS cveprod_oc, B.nomprod_oc AS nomprod_oc, " +
                "FORMAT(B.cantidad_oc, 'N2', 'es-mx') AS cantidad_oc, " +
                "FORMAT(B.precio_oc, 'N2', 'es-mx') AS precio_oc, FORMAT(B.importe_oc, 'N2', 'es-mx') AS importe_oc, B.conse " +
                "FROM tb_mstr_ordencompra A " +
                "JOIN tb_det_ordenescompra B ON A.numero_oc = B.numero_oc " +
                "JOIN tb_oc_cfdi C ON A.numero_oc = C.numero_oc " +
                "WHERE A.anticipada = '1' AND A.status_oc = 'A' AND " +
                "B.prod_clave = '" + lblClave.Text + "' AND A.cveprov_oc = '" + lblCveProv.Text + "' AND A.tipo_oc = '" + ((tipo == "NACIONAL") ? "N" : "E") + "' " +
                "AND C.recalculo = '0' ORDER BY A.fecha_oc DESC";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    dtgOrdenes.Rows.Add(reader1["numero_oc"], reader1["fecha_oc"], reader1["cveprod_oc"], reader1["nomprod_oc"], reader1["cantidad_oc"], reader1["precio_oc"], reader1["importe_oc"], reader1["conse"]);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
        }

        private void dtgOrdenes_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //cargar otras liquidaciones ligadas a esa orden de compra
            dtgLiquidaciones.DataSource = consultar_liquidaciones(dtgOrdenes.CurrentRow.Cells[0].Value.ToString(), lblCveProv.Text, tipo, lblClave.Text, dtgOrdenes.CurrentRow.Cells[1].Value.ToString());

            if (dtgLiquidaciones.Rows.Count > 0)
            {
                dtgLiquidaciones.Columns["Cajas"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dtgLiquidaciones.Columns["Precio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dtgLiquidaciones.Columns["Importe"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            bool found = false;
            if (validar_fecha_no_este_en_liquidacion(fecha1) == true)
            {
                found = true;
            }
            else
            {
                if (validar_fecha_no_este_en_liquidacion(fecha2) == true)
                {
                    found = true;
                }
            }
            if (found == true)
            {
                MessageBox.Show("Las fechas seleccionadas atraviesan otra liquidacion ya realizada anteriormente, verifique la información en las liquidaciones en el recuadro de abajo",
                    "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnAceptar.Enabled = false;
            }

            decimal liquidado = 0;
            foreach (DataGridViewRow gr in dtgLiquidaciones.Rows)
            {
                liquidado = liquidado + Convert.ToDecimal(gr.Cells["Cajas"].Value.ToString());
            }
            lblCantLiquidado.Text = liquidado.ToString("###,###,##0.00");

            decimal cant_oc = Convert.ToDecimal(dtgOrdenes.CurrentRow.Cells[4].Value.ToString());//cantidad de orden de compra
            decimal a_liquidar = Convert.ToDecimal(lblCantidad.Text);//cantidad a liquidar
            decimal sum = (a_liquidar + liquidado);//a liquidar + liquidado
            decimal diff = cant_oc - sum;
            if (liquidado == 0)//cuando no se ha liquidado nada
                lblDiferencia.Text = "0.00";
            else if (sum > cant_oc)
            {
                MessageBox.Show("La cantidad que quiere liquidar superará a la cantidad de la orden de compra: " + cant_oc.ToString("###,###,##0.00") + " cajas.\nLiquidado: " + liquidado.ToString("###,###,##0.00") +
                    " cajas\nCant. a Liquidar: " + a_liquidar.ToString("###,###,##0.00") + " cajas\nLiquidado + Cant. a Liquidar: " + sum.ToString("###,###,##0.00") + " cajas\n" +
                    "Diferencia de " + diff.ToString("###,###,##0.00") + " cajas", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                lblDiferencia.Text = diff.ToString("###,###,##0.00");

            if (liquidado >= cant_oc)
            {
                MessageBox.Show("Se deshabilitará el botón para cargar la orden de compra a la liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnAceptar.Enabled = false;
            }
            else
                btnAceptar.Enabled = true;
        }

        public class datos
        {
            private string _ordencompra;
            public string ordencompra
            {
                get { return _ordencompra; }
                set { _ordencompra = value; }
            }

            private string _recalculo;
            public string recalculo
            {
                get { return _recalculo; }
                set { _recalculo = value; }
            }

            private string _conse;
            public string conse
            {
                get { return _conse; }
                set { _conse = value; }
            }

            private string _terminar;
            public string terminar
            {
                get { return _terminar; }
                set { _terminar = value; }
            }
        }

        public class SharedDatos
        {
            public static datos DatosCell;
        }

        public DataTable consultar_liquidaciones(string oc, string prv, string tip, string prd, string fch)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Folio", typeof(string));
            dt.Columns.Add("Cajas", typeof(string));
            dt.Columns.Add("Precio", typeof(string));
            dt.Columns.Add("Importe", typeof(string));
            dt.Columns.Add("Fecha_1", typeof(string));
            dt.Columns.Add("Fecha_2", typeof(string));

            DataSet sd = new DataSet();
            SqlDataReader rdr1;
            SqlCommand cmnd2;
            SqlDataAdapter adptr;
            thisConnection.Open();
            //BUSCAR TODAS LAS LIQUIDACIONES DEL PROVEEDOR, PRODUCTO, LINEA Y FECHA MAYOR A LA ORDEN DE COMPRA
            string query = "";
            if (tipo == "EXPORTACION")
            {
                query = "SELECT A.liq_folio, A.liq_numoc1, A.liq_numoc2, A.liq_numoc3, A.liq_numoc4, A.uni_nac, A.liq_costo1, A.liq_imp_liq, A.uni_exp, A.liq_costo1e, A.liq_imp_liqe, " +
                "FORMAT(A.liq_fecha1, 'dd-MM-yyyy') AS liq_fecha1 , FORMAT(A.liq_fecha2, 'dd-MM-yyyy') AS liq_fecha2 " +
                "FROM tb_mstr_liquidacion A JOIN tb_det_liquidacion B ON A.liq_folio = B.liq_folio WHERE A.liq_provcve = '" + prv + "' " +
                "AND A.liq_prodcve = '" + prd + "' AND A.liq_fecha >= '" + fch + "' AND B.tipo_con = 'E' AND A.status = 'A' " +
                "GROUP BY A.liq_folio, A.liq_numoc1, A.liq_numoc2, A.liq_numoc3, A.liq_numoc4, A.uni_nac, A.liq_costo1, A.liq_imp_liq, A.uni_exp, A.liq_costo1e, A.liq_imp_liqe, liq_fecha1, liq_fecha2";
            }
            else
            {
                query = "SELECT A.liq_folio, A.liq_numoc1, A.liq_numoc2, A.liq_numoc3, A.liq_numoc4, A.uni_nac, A.liq_costo1, A.liq_imp_liq, A.uni_exp, A.liq_costo1e, A.liq_imp_liqe, " +
                "FORMAT(A.liq_fecha1, 'dd-MM-yyyy') AS liq_fecha1 , FORMAT(A.liq_fecha2, 'dd-MM-yyyy') AS liq_fecha2 " +
                "FROM tb_mstr_liquidacion A JOIN tb_det_liquidacion B ON A.liq_folio = B.liq_folio WHERE A.liq_provcve = '" + prv + "' " +
                "AND A.liq_prodcve = '" + prd + "' AND A.liq_fecha >= '" + fch + "' AND B.tipo_con = 'N' AND A.status = 'A' " +
                "GROUP BY A.liq_folio, A.liq_numoc1, A.liq_numoc2, A.liq_numoc3, A.liq_numoc4, A.uni_nac, A.liq_costo1, A.liq_imp_liq, A.uni_exp, A.liq_costo1e, A.liq_imp_liqe, liq_fecha1, liq_fecha2";
            }

            adptr = new SqlDataAdapter(query, thisConnection);
            adptr.Fill(sd, "liquidaciones");
            bool fnd = false;
            if (sd.Tables["liquidaciones"].Rows.Count > 0)
                fnd = true;


            if (fnd == true)
            {
                //BUSCAR LIQUIDACIONES DEPENDIENDO DEL TIPO DE ORDEN AGREGAR EL RENGLON AL DATATABLE
                DataRow r;
                foreach (DataRow rw in sd.Tables["liquidaciones"].Rows)
                {
                    if (rw["liq_numoc1"].ToString().Trim() == oc)
                    {
                        r = dt.NewRow();
                        if (tipo == "NACIONAL")
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_nac"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liq"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        else
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_exp"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1e"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liqe"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        dt.Rows.Add(r);
                    }
                    else if (rw["liq_numoc2"].ToString().Trim() == oc)
                    {
                        r = dt.NewRow();
                        if (tipo == "NACIONAL")
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_nac"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liq"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        else
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_exp"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1e"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liqe"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        dt.Rows.Add(r);
                    }
                    else if (rw["liq_numoc3"].ToString().Trim() == oc)
                    {
                        r = dt.NewRow();
                        if (tipo == "NACIONAL")
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_nac"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liq"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        else
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_exp"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1e"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liqe"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        dt.Rows.Add(r);
                    }
                    else if (rw["liq_numoc4"].ToString().Trim() == oc)
                    {
                        r = dt.NewRow();
                        if (tipo == "NACIONAL")
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_nac"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liq"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        else
                        {
                            r["Folio"] = rw["liq_folio"].ToString().Trim();
                            r["Cajas"] = Convert.ToDecimal(rw["uni_exp"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Precio"] = Convert.ToDecimal(rw["liq_costo1e"].ToString().Trim()).ToString("$###,###,##0.00");
                            r["Importe"] = Convert.ToDecimal(rw["liq_imp_liqe"].ToString().Trim()).ToString("###,###,##0.00");
                            r["Fecha_1"] = rw["liq_fecha1"].ToString().Trim();
                            r["Fecha_2"] = rw["liq_fecha2"].ToString().Trim();
                        }
                        dt.Rows.Add(r);
                    }
                }
            }
            thisConnection.Close();
            return dt;
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            datos val = new datos();
            string val1 = lblCantidad.Text;
            string val2 = dtgOrdenes.CurrentRow.Cells[4].Value.ToString();
            //if (Convert.ToDecimal(val1) != Convert.ToDecimal(val2))
            //{
            //    if (MessageBox.Show("La cantidad de cajas es diferente ¿Se hará el recalculo de la orden de compra?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            //        == System.Windows.Forms.DialogResult.Yes)
            //    {
            //        val.recalculo = "1";
            //    }
            //    else
            //        val.recalculo = "0";
            //}
            val.recalculo = "0";
            val.conse = dtgOrdenes.CurrentRow.Cells["conse"].Value.ToString();
            val.ordencompra = dtgOrdenes.CurrentRow.Cells[0].Value.ToString();
            SharedDatos.DatosCell = val;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        public bool validar_fecha_no_este_en_liquidacion(string fch)
        {
            bool fnd = false;
            foreach (DataGridViewRow rw in dtgLiquidaciones.Rows)
            {
                DateTime d1 = Convert.ToDateTime(rw.Cells["Fecha_1"].Value.ToString());
                DateTime d2 = Convert.ToDateTime(rw.Cells["Fecha_2"].Value.ToString());
                DateTime d3 = Convert.ToDateTime(fch);
                if ((d3 >= d1) && (d3 <= d2))
                    fnd = true;
                break;
            }
            return fnd;
        }

        private void articipadas_Load(object sender, EventArgs e)
        {

        }
    }
}
