using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Diagnostics;

namespace liquidacion
{
    public partial class Form1 : Form
    {
        //SqlConnection thisConnection = new SqlConnection("Data Source=GABIRA1\\SQL2005;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True");
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2, reader3;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;

        DataTable proveedor = new DataTable();
        DataTable lineas = new DataTable();
        DataTable productos = new DataTable();

        DataTable totalventas = new DataTable();

        DataTable tprod = new DataTable();

        DataTable datosliq = new DataTable();

        string cveprov = "";
        string nomprov = "";

        string cantidad = "";
        string tipo = "";

        string num_liq = "";
        string cant_orig = "";
        string cveprod = "";
        string nomprod = "";
        string cvelin = "";
        string nomlin = "";
        string f1 = "";
        string f2 = "";
        string neto_prod = "";

        string procedencia = "";

        DataTable dtequival = new DataTable();

        DataTable tmrepp1 = new DataTable();
        DataTable tmrepor1a = new DataTable();

        decimal pesotot = 0;

        DataTable dtkilos = new DataTable();
        DataTable dtrecibos = new DataTable();

        DataTable dtAnios = new DataTable();
        DataTable dtMeses = new DataTable();

        DataSet dsPrd = new DataSet();

        string tipo_reporte = "";
        string fech_ulti = "";

        bool nanana = true;

        DataTable dtLineas = new DataTable();

        public Form1()
        {
            InitializeComponent();

            //Process.Start(@"\\gabira1\liquidaciones\144794_NACIONAL.pdf");

            lblServidor.Text = Utilerias.Class1.ConnectionString;

            dtAnios.Columns.Add("anio", typeof(string));
            dtAnios.Columns.Add("desc", typeof(string));
            dtAnios.Rows.Add("00", "MES...");
            dtAnios.Rows.Add("01", "ENERO");
            dtAnios.Rows.Add("02", "FEBRERO");
            dtAnios.Rows.Add("03", "MARZO");
            dtAnios.Rows.Add("04", "ABRIL");
            dtAnios.Rows.Add("05", "MAYO");
            dtAnios.Rows.Add("06", "JUNIO");
            dtAnios.Rows.Add("07", "JULIO");
            dtAnios.Rows.Add("08", "AGOSTO");
            dtAnios.Rows.Add("09", "SEPTIEMBRE");
            dtAnios.Rows.Add("10", "OCTUBRE");
            dtAnios.Rows.Add("11", "NOVIEMBRE");
            dtAnios.Rows.Add("12", "DICIEMBRE");

            cmbAnio.DataSource = dtAnios;
            cmbAnio.DisplayMember = "desc";
            cmbAnio.ValueMember = "anio";

            dtMeses.Columns.Add("mes", typeof(string));
            dtMeses.Columns.Add("des", typeof(string));
            dtMeses.Rows.Add("2019", "2019");
            dtMeses.Rows.Add("2020", "2020");
            dtMeses.Rows.Add("2021", "2021");
            dtMeses.Rows.Add("2022", "2022");
            dtMeses.Rows.Add("2023", "2023");
            dtMeses.Rows.Add("2024", "2024");
            dtMeses.Rows.Add("2025", "2025");

            cmbAnio.DataSource = dtAnios;
            cmbAnio.DisplayMember = "desc";
            cmbAnio.ValueMember = "anio";

            cmbMes.DataSource = dtMeses;
            cmbMes.DisplayMember = "des";
            cmbMes.ValueMember = "mes";

            proveedor.Columns.Add("prov_clave", typeof(string));
            proveedor.Columns.Add("prov_nombre", typeof(string));

            lineas.Columns.Add("lin_clave", typeof(string));
            lineas.Columns.Add("lin_nombre", typeof(string));

            productos.Columns.Add("prod_clave", typeof(string));
            productos.Columns.Add("prod_nombre", typeof(string));
            productos.Columns.Add("lin_clave", typeof(string));

            totalventas.Columns.Add("prod_clave", typeof(string));
            totalventas.Columns.Add("nacional", typeof(string));
            totalventas.Columns.Add("exportacion", typeof(string));
            totalventas.Columns.Add("total", typeof(string));
            totalventas.Columns.Add("nalpor", typeof(string));
            totalventas.Columns.Add("exppor", typeof(string));

            tprod.Columns.Add("pro_clave", Type.GetType("System.String"));//0
            tprod.Columns.Add("pro_nombre", Type.GetType("System.String"));//1
            tprod.Columns.Add("lin_cve", Type.GetType("System.String"));//2
            tprod.Columns.Add("lin_nom", Type.GetType("System.String"));//3
            tprod.Columns.Add("unidades", Type.GetType("System.Decimal"));//4
            tprod.Columns.Add("uninac", Type.GetType("System.Decimal"));//5
            tprod.Columns.Add("uniexp", Type.GetType("System.Decimal"));//6
            tprod.Columns.Add("liquidado", Type.GetType("System.Boolean"));//7
            tprod.Columns.Add("num_liq", Type.GetType("System.String"));//8
            tprod.Columns.Add("neto", Type.GetType("System.Decimal"));//9
            tprod.Columns.Add("nal", Type.GetType("System.Decimal"));//10
            tprod.Columns.Add("ex", Type.GetType("System.Decimal"));//11
            tprod.Columns.Add("POR", Type.GetType("System.String"));//12
            tprod.Columns.Add("N", Type.GetType("System.String"));//13
            tprod.Columns.Add("E", Type.GetType("System.String"));//14
            tprod.Columns.Add("LN", Type.GetType("System.String"));//15
            tprod.Columns.Add("POR_IMP", Type.GetType("System.String"));//16

            datosliq.Columns.Add("liq_folio", Type.GetType("System.String"));//0
            datosliq.Columns.Add("uni_nac", Type.GetType("System.String"));//1
            datosliq.Columns.Add("uni_exp", Type.GetType("System.String"));//2

            tmrepp1.Columns.Add("folio", Type.GetType("System.String"));//0
            tmrepp1.Columns.Add("recibo", Type.GetType("System.String"));//1
            tmrepp1.Columns.Add("fecha", Type.GetType("System.DateTime"));//2
            tmrepp1.Columns.Add("linea", Type.GetType("System.String"));//3
            tmrepp1.Columns.Add("producto", Type.GetType("System.String"));//4
            tmrepp1.Columns.Add("producto2", Type.GetType("System.String"));//5
            tmrepp1.Columns.Add("tipoc", Type.GetType("System.String"));//6
            tmrepp1.Columns.Add("nombre", Type.GetType("System.String"));//7
            tmrepp1.Columns.Add("nombrelin", Type.GetType("System.String"));//8
            tmrepp1.Columns.Add("cantidad", Type.GetType("System.Decimal"));//9
            tmrepp1.Columns.Add("cantidad2", Type.GetType("System.Decimal"));//10
            tmrepp1.Columns.Add("util", Type.GetType("System.Decimal"));//11
            tmrepp1.Columns.Add("neto", Type.GetType("System.Decimal"));//12
            tmrepp1.Columns.Add("proveedor", Type.GetType("System.String"));//13
            tmrepp1.Columns.Add("nomprov", Type.GetType("System.String"));//14
            tmrepp1.Columns.Add("parcial", Type.GetType("System.Decimal"));//15
            tmrepp1.Columns.Add("cveran", Type.GetType("System.String"));//16
            tmrepp1.Columns.Add("nomran", Type.GetType("System.String"));//17
            tmrepp1.Columns.Add("cvetab", Type.GetType("System.String"));//18
            tmrepp1.Columns.Add("nomtab", Type.GetType("System.String"));//19
            tmrepp1.Columns.Add("tipo1", Type.GetType("System.String"));//20
            tmrepp1.Columns.Add("tipo2", Type.GetType("System.String"));//21

            tmrepor1a.Columns.Add("linea", Type.GetType("System.String"));
            tmrepor1a.Columns.Add("producto", Type.GetType("System.String"));
            tmrepor1a.Columns.Add("nombrelin", Type.GetType("System.String"));
            tmrepor1a.Columns.Add("nombre", Type.GetType("System.String"));
            tmrepor1a.Columns.Add("neto", Type.GetType("System.Decimal"));
            tmrepor1a.Columns.Add("cajas", Type.GetType("System.Decimal"));
            tmrepor1a.Columns.Add("POR", Type.GetType("System.String"));//12
            tmrepor1a.Columns.Add("N", Type.GetType("System.String"));//13
            tmrepor1a.Columns.Add("E", Type.GetType("System.String"));//14
            tmrepor1a.Columns.Add("LN", Type.GetType("System.String"));//15

            dtequival.Columns.Add("um_clave", typeof(string));
            dtequival.Columns.Add("compp_peso", typeof(string));
            dtequival.Columns.Add("prod_clave", typeof(string));
            dtequival.Columns.Add("um_equivalencia", typeof(string));

            dtkilos.Columns.Add("ordp_folio", typeof(string));
            dtkilos.Columns.Add("rmp_recibo", typeof(string));
            dtkilos.Columns.Add("rmp_tipo", typeof(string));
            dtkilos.Columns.Add("lin_clave", typeof(string));
            dtkilos.Columns.Add("prod_clave", typeof(string));
            dtkilos.Columns.Add("pesotot", typeof(string));

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            string filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Acceso 4.1 Liquidaciones");
                sw.Close();
            }

            //CARGA PROVEEDORES
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prov_clave, prov_nombre FROM tb_cat_proveedor ORDER BY prov_nombre";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = proveedor.NewRow();
                    rw["prov_clave"] = reader1.GetValue(0).ToString().Trim();
                    rw["prov_nombre"] = reader1.GetValue(1).ToString().Trim();
                    proveedor.Rows.Add(rw);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT lin_clave, lin_nombre FROM tb_cat_linea ORDER BY lin_nombre";
            reader1 = cmnd1.ExecuteReader();
            DataRow rx;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rx = lineas.NewRow();
                    rx["lin_clave"] = reader1.GetValue(0).ToString().Trim();
                    rx["lin_nombre"] = reader1.GetValue(1).ToString().Trim();
                    lineas.Rows.Add(rx);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prod_clave, prod_nombre, lin_clave FROM tb_cat_producto WHERE prod_nombre <> '' ORDER BY prod_nombre";
            reader1 = cmnd1.ExecuteReader();
            DataRow ry;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    ry = productos.NewRow();
                    ry["prod_clave"] = reader1.GetValue(0).ToString().Trim();
                    ry["prod_nombre"] = reader1.GetValue(1).ToString().Trim();
                    ry["lin_clave"] = reader1.GetValue(2).ToString().Trim();
                    productos.Rows.Add(ry);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT bloqueo FROM tb_liquidacion_bloqueo";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    lblBloqueo.Text = reader1["bloqueo"].ToString();
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            thisConnection.Close();

            foreach (DataRow rz in proveedor.Rows)
            {
                DDLProveedor.Items.Add(rz["prov_nombre"].ToString());
            }

            foreach (DataRow rz in lineas.Rows)
            {
                DDLLinea1.Items.Add(rz["lin_nombre"].ToString());
                DDLLinea2.Items.Add(rz["lin_nombre"].ToString());
            }

            //foreach (DataRow rz in productos.Rows)
            //{
            //    DDLEmpaques1.Items.Add(rz["prod_nombre"].ToString());
            //    DDLEmpaques2.Items.Add(rz["prod_nombre"].ToString());
            //}
        }

        private void btnGenera_Click(object sender, EventArgs e)
        {
            //string fecha_final = dtpFecha2.Text;
            //if (Convert.ToDateTime(fecha_final) >= DateTime.Now)
            //{
            //    MessageBox.Show("La fecha final del rango no debe ser mayor o igual a la fecha actual", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            ////ULTIMO CALCULO DE FLETES
            //string dia_inicial = "";
            //string dia_final = "";

            //dia_inicial = Convert.ToDateTime(dtpFecha1.Text).Day.ToString();
            //dia_final = Convert.ToDateTime(dtpFecha2.Text).Day.ToString();

            //Int32 dias = (Convert.ToDateTime(dtpFecha2.Text) - Convert.ToDateTime(dtpFecha1.Text)).Days;

            //if (Convert.ToDateTime(dtpFecha2.Text) == Convert.ToDateTime(dtpFecha1.Text))
            //{
            //    tipo_reporte = "S";
            //}
            //else
            //{
            //    if (dias == 6)//(Convert.ToInt32(dia_final) != Convert.ToInt32(DateTime.DaysInMonth(Convert.ToDateTime(dtpFecha2.Text).Year, Convert.ToDateTime(dtpFecha2.Text).Month)))
            //    {
            //        tipo_reporte = "S";
            //    }
            //    else
            //        tipo_reporte = "M";
            //}

            tipo_reporte = "M";
            lblUltimoCalculo.Text = ultima_fecha_calculada();

            if (Convert.ToDateTime(fech_ulti) < Convert.ToDateTime(dtpFecha2.Text))
            {
                MessageBox.Show("El calculo de los fletes no se ha realizado para las fechas solicitadas, favor de verificarlo", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                nanana = false;
            }
            else
                nanana = true;

            if (DDLTipo.SelectedIndex == 0)
            {
                producto_terminado();

                dtLineas = tprod.DefaultView.ToTable(true, "lin_cve");


            }
            if (DDLTipo.SelectedIndex == 1)
            {
                produccion();
            }


            DDLProveedor.Enabled = false;
            txtClaveProveedor.ReadOnly = true;
            //////REVISAR PRODUCTOS SUNSET
            ////decimal suma_caixas = 0;
            ////foreach (DataRow r in tprod.Select("pro_nombre like '%SUNSET%' AND uniexp > 0"))
            ////{
            ////    suma_caixas = suma_caixas + Convert.ToDecimal(r["uniexp"]);
            ////}
            ////MessageBox.Show("Caixas sunset: " + suma_caixas.ToString());

            ////decimal por_imp = 0;
            ////foreach (DataRow r in tprod.Select("pro_nombre like '%SUNSET%' AND uniexp > 0"))
            ////{
            ////    por_imp = (Convert.ToDecimal(r["uniexp"]) * 100) / suma_caixas;
            ////    r["POR_IMP"] = por_imp.ToString("##0.000");
            ////}

            ////foreach (DataRow r in tprod.Select("pro_nombre like '%SUNSET%' AND uniexp > 0"))
            ////{
            ////    foreach (DataGridViewRow gr in dtgLiquidacion.Rows)
            ////    {
            ////        if (r["pro_clave"].ToString() == gr.Cells["producto"].Value.ToString())
            ////        {
            ////            gr.Cells["por_imp"].Value= r["POR_IMP"];
            ////        }
            ////    }
            ////}

        }

        public void producto_terminado()
        {
            btnGenera.Enabled = false;
            label14.Visible = true;
            label14.Update();

            //VALIDACIONES
            #region validaciones
            if (DDLTipo.SelectedIndex == -1)
            {
                MessageBox.Show("Debe seleccionar el tipo de liquidación que se va a realizar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dtpFecha1.Text == "" || dtpFecha2.Text == "")
            {
                MessageBox.Show("Debe seleccionar un rango de fechas", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string prov = "";
            string lin1 = "";
            string lin2 = "";
            string pro1 = "";
            string pro2 = "";



            if (txtLinea1.Text == "")
                lin1 = "00000";
            else
                lin1 = txtLinea1.Text;
            if (txtLinea2.Text == "")
                lin2 = "99999";
            else
                lin2 = txtLinea2.Text;
            if (txtEmpaque1.Text == "")
                pro1 = "";
            else
                pro1 = txtEmpaque1.Text;
            if (txtEmpaque2.Text == "")
                pro2 = "ZZZZZZZZZZZ";
            else
                pro2 = txtEmpaque2.Text;
            #endregion

            thisConnection.Open();



            #region totalventas
            try
            {
                totalventas.Clear();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS nacional " +
                    "FROM tb_det_facturas DF, tb_mstr_facturas_nal F " +
                    "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + dtpFecha1.Text + "' and F.fcn_fecha <= '" + dtpFecha2.Text + "' " +
                    //"AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.fcn_lugar <> 'EXP' AND DF.fcn_tipo = F.fcn_lugar " +
                    "AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.um_clave = 'PESOS' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                    "GROUP BY DF.lin_clave, DF.prod_clave " +
                    "ORDER BY DF.lin_clave, DF.prod_clave";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw;
                while (reader1.Read())
                {
                    rw = totalventas.NewRow();
                    rw["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                    rw["nacional"] = reader1.GetValue(2).ToString().Trim();
                    rw["exportacion"] = "0";
                    totalventas.Rows.Add(rw);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenera.Enabled = true;
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            }

            try
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS nacional " +
                    "FROM tb_det_facturas DF, tb_mstr_facturas_nal F " +
                    "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + dtpFecha1.Text + "' and F.fcn_fecha <= '" + dtpFecha2.Text + "' " +
                    //"AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.fcn_lugar = 'EXP' AND DF.fcn_tipo = F.fcn_lugar " +
                    "AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.um_clave = 'USD' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                    "GROUP BY DF.lin_clave, DF.prod_clave " +
                    "ORDER BY DF.lin_clave, DF.prod_clave";
                reader1 = cmnd1.ExecuteReader();
                DataRow rw2;
                bool fnd = false;
                while (reader1.Read())
                {
                    foreach (DataRow rr in totalventas.Rows)
                    {
                        if (rr["prod_clave"].ToString() == reader1.GetValue(1).ToString().Trim())
                        {
                            fnd = true;
                            rr["exportacion"] = reader1.GetValue(2).ToString().Trim();
                        }
                    }
                    if (fnd == false)
                    {
                        rw2 = totalventas.NewRow();
                        rw2["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                        rw2["nacional"] = "0";
                        rw2["exportacion"] = reader1.GetValue(2).ToString().Trim();
                        totalventas.Rows.Add(rw2);
                    }
                    fnd = false;
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenera.Enabled = true;
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            }
            foreach (DataRow rr in totalventas.Rows)
            {
                decimal t = 0;
                t = (Convert.ToDecimal(rr["nacional"].ToString()) + Convert.ToDecimal(rr["exportacion"].ToString()));
                rr["total"] = (Convert.ToDecimal(rr["nacional"].ToString()) + Convert.ToDecimal(rr["exportacion"].ToString()));
                rr["nalpor"] = Math.Round(Convert.ToDecimal(rr["nacional"].ToString()) / t, 2).ToString();
                rr["exppor"] = Math.Round(Convert.ToDecimal(rr["exportacion"].ToString()) / t, 2).ToString();
            }
            #endregion

            #region generainventario

            try
            {
                tprod.Clear();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT P.lin_clave, H.prod_clave, SUM(H.hrp_num_unidades) AS hrp_num_unidades, SUM(H.hrp_peso_util) AS hrp_peso_util," +
                        "  H.hrp_numliq, M.prov_clave" +//H.hrp_liquidado,
                        "  " +
                        " FROM tb_hist_recepcion H, tb_mstr_recepcion_pt M, tb_cat_producto P" +
                        " WHERE H.hrp_recibo = M.rpt_recibo AND H.hrp_tipo_recepcion = 'PTC' AND (H.hrp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "'" +
                        " AND H.hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "') AND H.hrp_estatus <> 'C'" + //--AND M.prov_clave = '" + txtClaveProveedor.Text + "' " +
                        " AND (P.lin_clave >= '" + lin1 + "' AND P.lin_clave <= '" + lin2 + "') AND (H.prod_clave >= '" + pro1 + "' AND H.prod_clave <= '" + pro2 + "') AND" +
                        " H.hrp_situacion = 'CM' AND H.prod_clave = P.prod_clave" +
                        " AND (M.rpt_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' AND M.rpt_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "')" +
                        " group by H.prod_clave, P.lin_clave, H.hrp_tipo_recepcion, H.hrp_numliq, M.prov_clave" +//, H.hrp_liquidado
                        " ORDER BY H.prod_clave, H.hrp_tipo_recepcion";
                DataRow r;
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        if (reader1.GetValue(5).ToString().Trim() == txtClaveProveedor.Text)
                        {
                            r = tprod.NewRow();
                            r["pro_clave"] = reader1.GetValue(1).ToString().Trim();
                            r["pro_nombre"] = "";
                            r["lin_cve"] = reader1.GetValue(0).ToString().Trim();
                            r["lin_nom"] = "";
                            r["unidades"] = reader1.GetValue(2).ToString().Trim();
                            r["uninac"] = "0.00";
                            r["uniexp"] = "0.00";
                            r["liquidado"] = (reader1.GetValue(4).ToString().Trim() == "T") ? true : false;
                            r["num_liq"] = reader1.GetValue(4).ToString().Trim();
                            r["neto"] = reader1.GetValue(3).ToString().Trim();
                            r["nal"] = "0";
                            r["ex"] = "0";
                            tprod.Rows.Add(r);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No se encontro nungún dato para mostrar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    btnGenera.Enabled = true;
                    thisConnection.Close();
                    return;
                }
                cmnd1.Dispose();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenera.Enabled = true;
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
                return;
            }


            #endregion

            thisConnection.Close();

            cargadatosliquidaciones();

            foreach (DataRow rw in tprod.Rows)
            {
                rw["pro_nombre"] = nombreproducto(rw["pro_clave"].ToString());
                rw["lin_nom"] = nombrelinea(rw["lin_cve"].ToString());
            }

            DataView dw = tprod.DefaultView;
            dw.Sort = "lin_cve, pro_clave, liquidado";
            tprod = dw.ToTable();

            foreach (DataRow rw in tprod.Rows)
            {
                bool fnd = false;
                foreach (DataRow rs in totalventas.Select("prod_clave = '" + rw["pro_clave"].ToString() + "'"))
                {
                    fnd = true;
                    //operaciones
                    decimal nal = 0;
                    decimal exp = 0;
                    decimal rec = 0;
                    rec = Convert.ToDecimal(rw["unidades"].ToString());
                    nal = rec * Convert.ToDecimal(rs["nalpor"].ToString());
                    exp = rec * Convert.ToDecimal(rs["exppor"].ToString());

                    decimal val0 = Math.Round(nal, 0);
                    decimal val1 = nal - (val0);
                    decimal val2 = exp - (Math.Round(exp, 0));

                    if (val1 == Convert.ToDecimal("0.5") && val2 == Convert.ToDecimal("0.5"))
                    {
                        if (nal > exp)
                        {
                            exp = exp - val2;
                            nal = nal + val1;
                        }
                        else
                        {
                            exp = exp + val2;
                            nal = nal - val1;
                        }

                    }


                    rw["uninac"] = (nal == 0) ? "0.00" : Math.Round(nal, 0).ToString("###,###,##0.00");
                    rw["uniexp"] = (exp == 0) ? "0.00" : Math.Round(exp, 0).ToString("###,###,##0.00");

                    decimal total_por = rec;
                    decimal cien_por = Convert.ToDecimal("100");
                    decimal cajas_por = Convert.ToDecimal(rs["total"].ToString());
                    decimal por_total = ((cajas_por * cien_por) / total_por);

                    rw["POR"] = Math.Round(por_total, 2).ToString() + "%";
                    rw["N"] = Math.Round(Convert.ToDecimal(rs["nacional"].ToString()), 0).ToString();
                    rw["E"] = Math.Round(Convert.ToDecimal(rs["exportacion"].ToString()), 0).ToString();

                }
                if (fnd == false)
                {
                    rw["POR"] = "0%";
                    rw["N"] = "0";
                    rw["E"] = "0";
                }
            }

            decimal nacc = 0;
            decimal expor = 0;
            foreach (DataRow rw in tprod.Rows)
            {
                foreach (DataRow rs in datosliq.Select("liq_folio = '" + rw["num_liq"].ToString() + "'"))
                {
                    nacc = Convert.ToDecimal(rs["uni_nac"].ToString());
                    expor = Convert.ToDecimal(rs["uni_exp"].ToString());
                    break;
                }
                rw["nal"] = nacc.ToString();
                rw["ex"] = expor.ToString();
                nacc = 0;
                expor = 0;
            }
            //VERIFICACION DE SURTIDO POR RECEPCION Y NO POR VENTAS
            foreach (DataRow rr in tprod.Rows)
            {
                if (rr["num_liq"].ToString() != "")
                {
                    bool found = false;
                    foreach (DataRow rs in totalventas.Select("prod_clave = '" + rr["pro_clave"].ToString() + "'"))
                    {
                        found = true;
                    }
                    if (found == false)
                    {
                        rr["uninac"] = rr["nal"].ToString();
                        rr["uniexp"] = rr["ex"].ToString();
                    }
                    found = false;
                }
            }

            dtgLiquidacion.Rows.Clear();
            foreach (DataRow rw in tprod.Rows)
            {
                dtgLiquidacion.Rows.Add(rw["pro_clave"].ToString(), rw["pro_nombre"].ToString(), rw["lin_nom"].ToString(), Convert.ToDecimal(rw["unidades"].ToString()).ToString("###,###,##0.00"),
                    Convert.ToDecimal(rw["uninac"].ToString()).ToString("###,###,##0.00"), Convert.ToDecimal(rw["uniexp"].ToString()).ToString("###,###,##0.00"), rw["num_liq"].ToString(),
                    Convert.ToBoolean(rw["liquidado"].ToString()), Convert.ToDecimal(rw["nal"].ToString()).ToString("###,###,##0.00"), Convert.ToDecimal(rw["ex"].ToString()).ToString("###,###,##0.00"),
                    rw["POR"].ToString(), rw["N"].ToString(), rw["E"].ToString(), rw["lin_cve"].ToString());
            }



            colorear();

            string filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Consulta de liquidaciones de recepción pt");
                sw.Close();
            }

            label14.Visible = false;
            btnGenera.Enabled = true;
        }

        public void produccion()
        {
            btnGenera.Enabled = false;
            label14.Visible = true;
            label14.Update();



            //VALIDACIONES
            #region validaciones
            if (DDLTipo.SelectedIndex == -1)
            {
                MessageBox.Show("Debe seleccionar el tipo de liquidación que se va a realizar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dtpFecha1.Text == "" || dtpFecha2.Text == "")
            {
                MessageBox.Show("Debe seleccionar un rango de fechas", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string prov = "";
            string lin1 = "";
            string lin2 = "";
            string pro1 = "";
            string pro2 = "";



            if (txtLinea1.Text == "")
                lin1 = "00000";
            else
                lin1 = txtLinea1.Text;
            if (txtLinea2.Text == "")
                lin2 = "99999";
            else
                lin2 = txtLinea2.Text;
            if (txtEmpaque1.Text == "")
                pro1 = "";
            else
                pro1 = txtEmpaque1.Text;
            if (txtEmpaque2.Text == "")
                pro2 = "ZZZZZZZZZZZ";
            else
                pro2 = txtEmpaque2.Text;

            tprod.Clear();
            tmrepor1a.Clear();
            dtequival.Clear();
            tmrepp1.Clear();

            #endregion

            thisConnection.Open();

            string qry = "";

            #region pesotot


            //string qry = "SELECT P.ordp_folio, P.rmp_recibo, P.rmp_tipo, P.lin_clave, P.prod_clave, H.hrp_clase1, H.hrp_num_unidades, P.podp_cantidad " +
            //    "FROM tb_det_prod_odp P, tb_hist_recepcion H " +
            //    "WHERE P.rmp_recibo = H.hrp_recibo AND P.lin_clave = H.lin_clave AND P.rmp_tipo = H.hrp_tipo_recepcion AND P.lin_clave = H.lin_clave " +
            //    "AND P.prod_clave = H.prod_clave and H.hrp_clase1 > 0 AND H.hrp_fecha >= '01/09/2020' and H.hrp_fecha <= '28/09/2020' " +
            //    "order by rmp_recibo, P.rmp_tipo, P.prod_clave, P.lin_clave";
            //SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
            //adapter.Fill(dsPrd, "dtKilos");


            //cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = qry;
            //reader1 = cmnd1.ExecuteReader();
            //DataRow ra;
            //if (reader1.HasRows)
            //{
            //    while (reader1.Read())
            //    {
            //        ra = dtkilos.NewRow();
            //        ra["ordp_folio"] = reader1.GetValue(0).ToString().Trim();
            //        ra["rmp_recibo"] = reader1.GetValue(1).ToString().Trim();
            //        ra["rmp_tipo"] = reader1.GetValue(2).ToString().Trim();
            //        ra["lin_clave"] = reader1.GetValue(3).ToString().Trim();
            //        ra["prod_clave"] = reader1.GetValue(4).ToString().Trim();
            //        ra["pesotot"] = reader1.GetValue(5).ToString().Trim();
            //        dtkilos.Rows.Add(ra);
            //    }
            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();
            #endregion

            SqlDataAdapter adapter;
            #region totalventas
            try
            {
                qry = "SELECT /*DF.lin_clave,*/ DF.prod_clave, SUM(DF.fcn_num_unidades) AS nacional " +
                    "FROM tb_det_facturas DF, tb_mstr_facturas_nal F " +
                    "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + dtpFecha1.Text + "' and F.fcn_fecha <= '" + dtpFecha2.Text + "' " +
                    //"AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.fcn_lugar <> 'EXP' AND DF.fcn_tipo = F.fcn_lugar " +
                    "AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.um_clave = 'PESOS' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                    "GROUP BY /*DF.lin_clave,*/ DF.prod_clave " +
                    "ORDER BY /*DF.lin_clave,*/ DF.prod_clave";

                totalventas.Clear();

                adapter = new SqlDataAdapter(qry, thisConnection);
                adapter.Fill(totalventas);

                foreach (DataRow re in totalventas.Rows)
                {
                    re["exportacion"] = "0";
                }

                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = qry;
                //reader1 = cmnd1.ExecuteReader();
                //DataRow rw;
                //while (reader1.Read())
                //{
                //    rw = totalventas.NewRow();
                //    rw["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                //    rw["nacional"] = reader1.GetValue(2).ToString().Trim();
                //    rw["exportacion"] = "0";
                //    totalventas.Rows.Add(rw);
                //}
                //reader1.Close();
                //reader1.Dispose();
                //cmnd1.Dispose();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenera.Enabled = true;
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            }

            try
            {
                qry = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion " +
                    "FROM tb_det_facturas DF, tb_mstr_facturas_nal F " +
                    "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + dtpFecha1.Text + "' and F.fcn_fecha <= '" + dtpFecha2.Text + "' " +
                    //"AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.fcn_lugar = 'EXP' AND DF.fcn_tipo = F.fcn_lugar " +
                    "AND DF.lin_clave >= '" + txtLinea1.Text + "' AND DF.lin_clave <= '" + txtLinea2.Text + "' AND F.um_clave = 'USD' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                    "GROUP BY DF.lin_clave, DF.prod_clave " +
                    "ORDER BY DF.lin_clave, DF.prod_clave";
                adapter = new SqlDataAdapter(qry, thisConnection);
                adapter.Fill(dsPrd, "ventasExportacion");

                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = qry;
                //reader1 = cmnd1.ExecuteReader();
                DataRow rw2;
                bool fnd = false;
                //while (reader1.Read())
                //{
                foreach (DataRow rr in dsPrd.Tables["ventasExportacion"].Rows)
                {
                    foreach (DataRow r1 in totalventas.Select("prod_clave = '" + rr["prod_clave"] + "'"))
                    {
                        fnd = true;
                    }
                    if (fnd == false)
                    {
                        rw2 = totalventas.NewRow();
                        rw2["prod_clave"] = rr["prod_clave"];
                        rw2["nacional"] = "0";
                        rw2["exportacion"] = rr["exportacion"];
                        totalventas.Rows.Add(rw2);
                    }
                    fnd = false;
                }

                //foreach (DataRow rr in totalventas.Rows)
                //{
                //    if (rr["prod_clave"].ToString() == reader1.GetValue(1).ToString().Trim())
                //    {
                //        fnd = true;
                //        rr["exportacion"] = reader1.GetValue(2).ToString().Trim();
                //    }
                //}
                //if (fnd == false)
                //{
                //    rw2 = totalventas.NewRow();
                //    rw2["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                //    rw2["nacional"] = "0";
                //    rw2["exportacion"] = reader1.GetValue(2).ToString().Trim();
                //    totalventas.Rows.Add(rw2);
                //}

                //}
                //reader1.Close();
                //reader1.Dispose();
                //cmnd1.Dispose();
            }
            catch (SqlException ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                btnGenera.Enabled = true;
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            }
            foreach (DataRow rr in totalventas.Rows)
            {
                decimal t = 0;
                t = (Convert.ToDecimal(rr["nacional"].ToString()) + Convert.ToDecimal(rr["exportacion"].ToString()));
                rr["total"] = (Convert.ToDecimal(rr["nacional"].ToString()) + Convert.ToDecimal(rr["exportacion"].ToString()));
                rr["nalpor"] = Math.Round(Convert.ToDecimal(rr["nacional"].ToString()) / t, 2).ToString();
                rr["exppor"] = Math.Round(Convert.ToDecimal(rr["exportacion"].ToString()) / t, 2).ToString();
            }
            #endregion

            #region generainventario

            //try
            //{
            //    tprod.Clear();
            //    cmnd1 = thisConnection.CreateCommand();
            //    cmnd1.CommandText = "SELECT H.lin_clave, H.prod_clave, SUM(H.hrp_num_unidades) AS hrp_num_unidades, SUM(H.hrp_peso_util) AS hrp_peso_util," +
            //            " H.hrp_liquidado, H.hrp_numliq, M.prov_clave" +
            //            "  " +
            //            " FROM tb_hist_recepcion H, tb_mstr_recepcion_pt M " +
            //            " WHERE H.hrp_recibo = M.rpt_recibo AND H.hrp_tipo_recepcion = 'PTC' AND (H.hrp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "'" +
            //            " AND H.hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "') AND H.hrp_estatus <> 'C'" + //--AND M.prov_clave = '" + txtClaveProveedor.Text + "' " +
            //            " AND (H.lin_clave >= '" + lin1 + "' AND H.lin_clave <= '" + lin2 + "') AND (H.prod_clave >= '" + pro1 + "' AND H.prod_clave <= '" + pro2 + "') AND" +
            //            " H.hrp_situacion = 'CM' " +
            //            " group by H.prod_clave, H.lin_clave, H.hrp_tipo_recepcion, H.hrp_numliq, H.hrp_liquidado, M.prov_clave" +
            //            " ORDER BY H.prod_clave, H.hrp_tipo_recepcion";
            //    DataRow r;
            //    reader1 = cmnd1.ExecuteReader();
            //    if (reader1.HasRows)
            //    {
            //        while (reader1.Read())
            //        {
            //            if (reader1.GetValue(6).ToString().Trim() == txtClaveProveedor.Text)
            //            {
            //                r = tprod.NewRow();
            //                r["pro_clave"] = reader1.GetValue(1).ToString().Trim();
            //                r["pro_nombre"] = "";
            //                r["lin_cve"] = reader1.GetValue(0).ToString().Trim();
            //                r["lin_nom"] = "";
            //                r["unidades"] = reader1.GetValue(2).ToString().Trim();
            //                r["uninac"] = "0.00";
            //                r["uniexp"] = "0.00";
            //                r["liquidado"] = (reader1.GetValue(4).ToString().Trim() == "T") ? true : false;
            //                r["num_liq"] = reader1.GetValue(5).ToString().Trim();
            //                r["neto"] = reader1.GetValue(3).ToString().Trim();
            //                r["nal"] = "0";
            //                r["ex"] = "0";
            //                tprod.Rows.Add(r);
            //            }
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show("No se encontro nungún dato para mostrar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        btnGenera.Enabled = true;
            //        thisConnection.Close();
            //        return;
            //    }
            //    cmnd1.Dispose();
            //}
            //catch (SqlException ex)
            //{
            //    MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    btnGenera.Enabled = true;
            //    if (thisConnection.State == ConnectionState.Open)
            //        thisConnection.Close();
            //    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
            //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            //    return;
            //}


            #endregion

            //Correcto 29/09/2020 11:59 a.m.

            thisConnection.Close();

            DataTable dtfinodp = new DataTable();
            dtfinodp.Columns.Add("ordp_folio", typeof(string));
            dtfinodp.Columns.Add("prod_clave", typeof(string));
            dtfinodp.Columns.Add("prod_nombre", typeof(string));
            dtfinodp.Columns.Add("lin_clave", typeof(string));
            dtfinodp.Columns.Add("lin_nombre", typeof(string));
            dtfinodp.Columns.Add("fodp_unidades", typeof(string));
            dtfinodp.Columns.Add("prod_peso_var", typeof(string));
            dtfinodp.Columns.Add("prod_clasificacion", typeof(string));
            DataTable dtdetord = new DataTable();
            DataTable compprod = new DataTable();
            DataTable recepcionmp = new DataTable();
            DataTable provrchtbl = new DataTable();
            DataTable comprodum = new DataTable();
            DataTable histo = new DataTable();
            DataTable hisprod = new DataTable();

            string ln1 = "";
            string ln2 = "";

            try
            {

                string pr1 = "";
                string pr2 = "";

                if (txtLinea1.Text == "")
                    ln1 = "00000";
                else
                    ln1 = txtLinea1.Text;
                if (txtLinea2.Text == "")
                    ln2 = "99999";
                else
                    ln2 = txtLinea2.Text;

                if (txtEmpaque1.Text == "")
                    pr1 = "";
                else
                    pr1 = txtEmpaque1.Text;
                if (txtEmpaque2.Text == "")
                    pr2 = "zzzzzzzzzz";
                else
                    pr2 = txtEmpaque2.Text;

                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT F.ordp_folio, F.prod_clave, P.prod_nombre, F.lin_clave, L.lin_nombre, F.fodp_unidades, F.prod_peso_var, P.prod_clasificacion FROM tb_det_final_odp F, tb_cat_linea L, tb_cat_producto P WHERE F.lin_clave = L.lin_clave AND F.prod_clave = P.prod_clave" +
                //            " AND (F.lin_clave >= '" + ln1 + "' AND F.lin_clave <= '" + ln2 + "') AND (F.prod_clave >= '" + pr1 + "' AND F.prod_clave <= '" + pr2 + "') order by F.ordp_folio";//F.ordp_folio DESC
                cmnd1.CommandText = "SELECT F.ordp_folio, F.prod_clave, P.prod_nombre, F.lin_clave, L.lin_nombre, F.fodp_unidades, F.prod_peso_var, P.prod_clasificacion " +
                    "FROM tb_det_final_odp F join tb_cat_linea L ON F.lin_clave = L.lin_clave " +
                    "JOIN tb_cat_producto P ON F.prod_clave = P.prod_clave " +
                    "JOIN tb_mstr_ordenes_prod A ON F.ordp_folio = A.ordp_folio " +
                    "WHERE (F.lin_clave >= '00000' AND F.lin_clave <= 'ZZZZZ') AND (F.prod_clave >= '0000000000' AND F.prod_clave <= 'ZZZZZZZZZZ') " +
                    "AND A.ordp_fecha >=  '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' and A.ordp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "' order by F.ordp_folio";
                reader1 = cmnd1.ExecuteReader();
                DataRow rfinodp;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rfinodp = dtfinodp.NewRow();
                        rfinodp["ordp_folio"] = reader1.GetValue(0).ToString().Trim();
                        rfinodp["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                        rfinodp["prod_nombre"] = reader1.GetValue(2).ToString().Trim();
                        rfinodp["lin_clave"] = reader1.GetValue(3).ToString().Trim();
                        rfinodp["lin_nombre"] = reader1.GetValue(4).ToString().Trim();
                        rfinodp["fodp_unidades"] = reader1.GetValue(5).ToString().Trim();
                        rfinodp["prod_peso_var"] = reader1.GetValue(6).ToString().Trim();
                        rfinodp["prod_clasificacion"] = reader1.GetValue(7).ToString().Trim();
                        dtfinodp.Rows.Add(rfinodp);
                    }
                }
                reader1.Close();
                reader1.Dispose();


                DataRow rdetord;
                dtdetord.Columns.Add("ordp_folio", typeof(string));
                dtdetord.Columns.Add("rmp_recibo", typeof(string));
                dtdetord.Columns.Add("lin_clave", typeof(string));
                dtdetord.Columns.Add("prod_clave", typeof(string));
                dtdetord.Columns.Add("rmp_tipo", typeof(string));
                dtdetord.Columns.Add("podp_cantidad", typeof(string));
                cmnd1.CommandText = "SELECT B.ordp_folio, B.rmp_recibo, B.lin_clave, B.prod_clave, B.rmp_tipo, B.podp_cantidad FROM tb_mstr_ordenes_prod A, tb_det_prod_odp B" +
                    " WHERE A.ordp_folio = B.ordp_folio AND (A.ordp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' AND A.ordp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "')" +
                    " ORDER BY B.rmp_recibo DESC";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rdetord = dtdetord.NewRow();
                        rdetord["ordp_folio"] = reader1.GetValue(0).ToString().Trim();
                        rdetord["rmp_recibo"] = reader1.GetValue(1).ToString().Trim();
                        rdetord["lin_clave"] = reader1.GetValue(2).ToString().Trim();
                        rdetord["prod_clave"] = reader1.GetValue(3).ToString().Trim();
                        rdetord["rmp_tipo"] = reader1.GetValue(4).ToString().Trim();
                        rdetord["podp_cantidad"] = reader1.GetValue(5).ToString().Trim();
                        dtdetord.Rows.Add(rdetord);
                    }
                }
                reader1.Close();
                reader1.Dispose();

                provrchtbl.Columns.Add("rmp_recibo", typeof(string));
                provrchtbl.Columns.Add("rch_clave", typeof(string));
                provrchtbl.Columns.Add("tbl_clave", typeof(string));
                provrchtbl.Columns.Add("prov_clave", typeof(string));
                provrchtbl.Columns.Add("rch_nombre", typeof(string));
                provrchtbl.Columns.Add("tbl_nombre", typeof(string));
                provrchtbl.Columns.Add("prov_nombre", typeof(string));
                DataRow rowprov;
                cmnd1.CommandText = "SELECT M.rmp_recibo, M.rch_clave, M.tbl_clave, M.prov_clave, R.rch_nombre, T.tbl_nombre, P.prov_nombre" +
                    " FROM tb_mstr_recepcion_mp M, tb_cat_tablas T, tb_cat_proveedor P, tb_cat_ranchos R" +
                    " WHERE M.prov_clave = P.prov_clave AND M.prov_clave = R.prov_clave and M.rch_clave = R.rch_clave AND M.prov_clave = T.prov_clave and" +
                    " M.rch_clave = T.rch_clave AND M.tbl_clave = T.tbl_clave " +
                    "/*AND M.rmp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' and M.rmp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "'*/";
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rowprov = provrchtbl.NewRow();
                        rowprov["rmp_recibo"] = reader1.GetValue(0).ToString().Trim();
                        rowprov["rch_clave"] = reader1.GetValue(1).ToString().Trim();
                        rowprov["tbl_clave"] = reader1.GetValue(2).ToString().Trim();
                        rowprov["prov_clave"] = reader1.GetValue(3).ToString().Trim();
                        rowprov["rch_nombre"] = reader1.GetValue(4).ToString().Trim();
                        rowprov["tbl_nombre"] = reader1.GetValue(5).ToString().Trim();
                        rowprov["prov_nombre"] = reader1.GetValue(6).ToString().Trim();
                        provrchtbl.Rows.Add(rowprov);
                    }
                }
                reader1.Close();
                reader1.Dispose();





                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "select A.um_clave, A.compp_peso, A.prod_clave, B.um_equivalencia FROM tb_mstr_comp_prod A, tb_cat_unidad B" +
                    " WHERE A.um_clave = B.um_clave AND (A.lin_clave >= '" + ln1 + "' AND A.lin_clave <= '" + ln2 + "')";
                reader1 = cmnd1.ExecuteReader();
                DataRow rwe;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rwe = dtequival.NewRow();
                        rwe["um_clave"] = reader1.GetValue(0).ToString().Trim();
                        rwe["compp_peso"] = reader1.GetValue(1).ToString().Trim();
                        rwe["prod_clave"] = reader1.GetValue(2).ToString().Trim();
                        rwe["um_equivalencia"] = reader1.GetValue(3).ToString().Trim();
                        dtequival.Rows.Add(rwe);
                    }

                }
                reader1.Close();
                reader1.Dispose();

                thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT C.um_clave, C.compp_peso, U.um_equivalencia, C.prod_clave FROM tb_mstr_comp_prod C, tb_cat_unidad U WHERE C.um_clave = U.um_clave ORDER BY C.prod_clave";
                reader1 = cmnd1.ExecuteReader();
                comprodum.Columns.Add("um_clave", typeof(string));
                comprodum.Columns.Add("compp_peso", typeof(string));
                comprodum.Columns.Add("um_equivalencia", typeof(string));
                comprodum.Columns.Add("prod_clave", typeof(string));
                DataRow rum;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rum = comprodum.NewRow();
                        rum["um_clave"] = reader1.GetValue(0).ToString().Trim();
                        rum["compp_peso"] = reader1.GetValue(1).ToString().Trim();
                        rum["um_equivalencia"] = reader1.GetValue(2).ToString().Trim();
                        rum["prod_clave"] = reader1.GetValue(3).ToString().Trim();
                        comprodum.Rows.Add(rum);
                    }
                }
                reader1.Close();
                reader1.Dispose();

                thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT hrp_recibo, hrp_tipo_recepcion, lin_clave, prod_clave, hrp_clase1, hrp_num_unidades" +
                    " FROM tb_hist_recepcion --where (hrp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' and hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "')";
                reader1 = cmnd1.ExecuteReader();
                histo.Columns.Add("hrp_recibo", typeof(string));
                histo.Columns.Add("hrp_tipo_recepcion", typeof(string));
                histo.Columns.Add("lin_clave", typeof(string));
                histo.Columns.Add("prod_clave", typeof(string));
                histo.Columns.Add("hrp_clase1", typeof(string));
                histo.Columns.Add("hrp_num_unidades", typeof(string));
                DataRow rh;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rh = histo.NewRow();
                        rh["hrp_recibo"] = reader1.GetValue(0).ToString().Trim();
                        rh["hrp_tipo_recepcion"] = reader1.GetValue(1).ToString().Trim();
                        rh["lin_clave"] = reader1.GetValue(2).ToString().Trim();
                        rh["prod_clave"] = reader1.GetValue(3).ToString().Trim();
                        rh["hrp_clase1"] = reader1.GetValue(4).ToString().Trim();
                        rh["hrp_num_unidades"] = reader1.GetValue(5).ToString().Trim();
                        histo.Rows.Add(rh);
                    }
                }
                reader1.Close();
                reader1.Dispose();

                thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT P.ordp_folio, P.rmp_recibo, P.rmp_tipo, P.lin_clave, P.prod_clave, " +
                    "(CASE " +
                        "when P.lin_clave = '08' then H.hrp_clase2 " +
                        "else H.hrp_clase1 END) AS hrp_clase1 " +
                    "/*H.hrp_clase1*/, H.hrp_num_unidades, P.podp_cantidad" +
                    " FROM tb_det_prod_odp P, tb_hist_recepcion H" +
                    " WHERE P.rmp_recibo = H.hrp_recibo AND P.lin_clave = H.lin_clave AND P.rmp_tipo = H.hrp_tipo_recepcion AND P.lin_clave = H.lin_clave" +
                    " AND P.prod_clave = H.prod_clave" +
                    " /*AND (H.hrp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' AND H.hrp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "')*/" +
                    " order by rmp_recibo, P.rmp_tipo, P.prod_clave, P.lin_clave";
                reader1 = cmnd1.ExecuteReader();
                hisprod.Columns.Add("ordp_folio", typeof(string));
                hisprod.Columns.Add("rmp_recibo", typeof(string));
                hisprod.Columns.Add("rmp_tipo", typeof(string));
                hisprod.Columns.Add("lin_clave", typeof(string));
                hisprod.Columns.Add("prod_clave", typeof(string));
                hisprod.Columns.Add("hrp_clase1", typeof(string));
                hisprod.Columns.Add("hrp_num_unidades", typeof(string));
                hisprod.Columns.Add("podp_cantidad", typeof(string));
                DataRow rhp;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rh = hisprod.NewRow();
                        rh["ordp_folio"] = reader1.GetValue(0).ToString().Trim();
                        rh["rmp_recibo"] = reader1.GetValue(1).ToString().Trim();
                        rh["rmp_tipo"] = reader1.GetValue(2).ToString().Trim();
                        rh["lin_clave"] = reader1.GetValue(3).ToString().Trim();
                        rh["prod_clave"] = reader1.GetValue(4).ToString().Trim();
                        rh["hrp_clase1"] = reader1.GetValue(5).ToString().Trim();
                        rh["hrp_num_unidades"] = reader1.GetValue(6).ToString().Trim();
                        rh["podp_cantidad"] = reader1.GetValue(7).ToString().Trim();
                        hisprod.Rows.Add(rh);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());

                return;
            }

            decimal pesocom = 0;
            string tipo = "";
            decimal var_caj_recibo = 0;
            decimal peso3 = 0;

            string var_chr_prod_nombre = "";
            string var_chr_tipoc = "";

            string nom_lin = "";

            int conta = 0;
            string var_chr_ordp_folio = "";
            string var_date_ordp_fecha = "";
            string var_chr_lin_clave = "";
            string var_chr_prod_clave = "";
            decimal var_dec_fodp_unidades = 0;
            string var_chr_rmp_recibo = "";
            string var_lin_clave = "";
            string var_prod_clave = "";
            string var_chr_um_clave = "";
            decimal var_dec_um_equivalencia = 0;
            decimal var_dec_peso_total = 0;

            tmrepp1.Clear();
            tmrepor1a.Clear();
            tprod.Clear();
            dtgLiquidacion.Rows.Clear();
            try
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT ordp_folio, ordp_fecha FROM tb_mstr_ordenes_prod WHERE (ordp_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' AND " +
                    "ordp_fecha <= '" + Convert.ToDateTime(dtpFecha2.Text).ToShortDateString() + "') AND ordp_estatus <> 'C'"; // 
                reader1 = cmnd1.ExecuteReader();
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        var_chr_ordp_folio = reader1.GetValue(0).ToString().Trim();
                        var_date_ordp_fecha = reader1.GetValue(1).ToString().Trim();


                        foreach (DataRow r in dtfinodp.Select("ordp_folio = '" + var_chr_ordp_folio + "'"))
                        {
                            var_chr_lin_clave = r[3].ToString();
                            var_chr_prod_clave = r[1].ToString();
                            var_dec_fodp_unidades = Convert.ToDecimal(r[5].ToString());
                            pesocom = Convert.ToDecimal(r[6].ToString());
                            nom_lin = r[4].ToString();
                            var_chr_prod_nombre = r[2].ToString();
                            var_chr_tipoc = r[7].ToString();

                            foreach (DataRow r1 in dtdetord.Select("ordp_folio = '" + var_chr_ordp_folio + "'"))
                            {
                                var_chr_rmp_recibo = r1[1].ToString();
                                var_lin_clave = r1[2].ToString();
                                var_prod_clave = r1[3].ToString();
                                tipo = r1[4].ToString();
                                var_caj_recibo = Convert.ToDecimal(r1[5].ToString());

                                if (pesocom == 0)
                                {
                                    foreach (DataRow rw in comprodum.Select("prod_clave = '" + var_chr_prod_clave + "'"))
                                    {
                                        var_chr_um_clave = rw["um_clave"].ToString();
                                        var_dec_um_equivalencia = Convert.ToDecimal(rw["um_equivalencia"].ToString());
                                        pesocom = pesocom + (Convert.ToDecimal(rw["compp_peso"].ToString()) * var_dec_um_equivalencia);
                                    }
                                }
                                string cver = "";
                                string cvet = "";
                                string nomr = "";
                                string nomt = "";
                                string cveprov = "";
                                string t1 = "";
                                string t2 = "";
                                string nomprove = "";
                                if (tipo == "REM")
                                {
                                    cver = "";
                                    cvet = "";
                                    nomr = "";
                                    nomt = "";
                                    cveprov = "";
                                    t1 = "REM";
                                    t2 = "REM";
                                }
                                else
                                {

                                    if (var_chr_lin_clave == "08")
                                    {
                                        #region esparrago
                                        t1 = "PRO";
                                        t2 = "ESP";
                                        cmnd1 = thisConnection.CreateCommand();
                                        cmnd1.CommandText = "SELECT rch_clave, tbl_clave, prov_clave FROM tb_mstr_recepcion_esparrago WHERE rmp_recibo = '" + var_chr_rmp_recibo + "'";
                                        reader3 = cmnd1.ExecuteReader();
                                        if (reader3.HasRows)
                                        {
                                            while (reader3.Read())
                                            {
                                                cver = reader3.GetValue(0).ToString().Trim();
                                                cvet = reader3.GetValue(1).ToString().Trim();
                                                cveprov = reader3.GetValue(2).ToString().Trim();

                                                cmnd2 = thisConnection.CreateCommand();
                                                cmnd2.CommandText = "SELECT hrp_clase2, hrp_num_unidades FROM tb_hist_recepcion WHERE hrp_recibo = '" + var_chr_rmp_recibo + "'" +
                                                    " AND hrp_tipo_recepcion = '" + tipo + "' AND lin_clave = '" + var_lin_clave + "' AND prod_clave = '" + var_prod_clave + "'";
                                                reader2 = cmnd2.ExecuteReader();
                                                if (reader2.HasRows)
                                                {
                                                    while (reader2.Read())
                                                    {
                                                        peso3 = (reader2.GetDecimal(0) / reader2.GetDecimal(1)) * var_caj_recibo;
                                                    }
                                                }
                                                else
                                                    peso3 = 0;
                                                reader2.Close();
                                                reader2.Dispose();

                                                cmnd2.CommandText = "SELECT prov_nombre FROM tb_cat_proveedor WHERE prov_clave = '" + cveprov + "'";
                                                reader2 = cmnd2.ExecuteReader();
                                                while (reader2.Read())
                                                {
                                                    nomprove = reader2.GetValue(0).ToString().Trim();
                                                }
                                                reader2.Close();
                                                reader2.Dispose();

                                                cmnd2.CommandText = "SELECT rch_nombre FROM tb_cat_ranchos WHERE prov_clave = '" + cveprov + "' AND rch_clave = '" + cver + "'";
                                                reader2 = cmnd2.ExecuteReader();
                                                while (reader2.Read())
                                                {
                                                    nomr = reader2.GetValue(0).ToString().Trim();
                                                }
                                                reader2.Close();
                                                reader2.Dispose();

                                                cmnd2.CommandText = "SELECT tbl_nombre FROM tb_cat_tablas WHERE prov_clave = '" + cveprov + "' AND rch_clave = '" + cver + "' AND tbl_clave = '" + cvet + "'";
                                                reader2 = cmnd2.ExecuteReader();
                                                while (reader2.Read())
                                                {
                                                    nomt = reader2.GetValue(0).ToString().Trim();
                                                }
                                                reader2.Close();
                                                reader2.Dispose();
                                            }
                                        }
                                        else
                                        {
                                            cver = "";
                                            cvet = "";
                                            nomr = "";
                                            nomt = "";
                                            cveprov = "";
                                        }
                                        reader3.Close();
                                        reader3.Dispose();

                                        if (String.Compare(cveprov, txtClaveProveedor.Text, true) == 0)
                                        {
                                            DataRow drw = tmrepp1.NewRow();
                                            drw["folio"] = var_chr_ordp_folio;
                                            drw["recibo"] = var_chr_rmp_recibo;
                                            drw["linea"] = var_chr_lin_clave;
                                            drw["fecha"] = var_date_ordp_fecha;
                                            drw["producto"] = var_chr_prod_clave;
                                            drw["cantidad"] = var_dec_fodp_unidades;


                                            var_dec_peso_total = 0; // fn_peso_total(var_chr_lin_clave, var_chr_prod_clave);
                                            foreach (DataRow rw in comprodum.Select("prod_clave = '" + var_chr_prod_clave + "'"))
                                            {
                                                var_chr_um_clave = rw["um_clave"].ToString();
                                                var_dec_um_equivalencia = Convert.ToDecimal(rw["um_equivalencia"].ToString());
                                                var_dec_peso_total = var_dec_peso_total + (Convert.ToDecimal(rw["compp_peso"].ToString()) * var_dec_um_equivalencia);
                                            }

                                            drw["nombre"] = var_chr_prod_nombre;
                                            drw["tipoc"] = var_chr_tipoc;
                                            drw["neto"] = pesocom;
                                            drw["proveedor"] = cveprov;
                                            drw["nomprov"] = nomprov;
                                            drw["nombrelin"] = nom_lin;
                                            drw["cveran"] = cver;
                                            drw["nomran"] = nomr;
                                            drw["cvetab"] = cvet;
                                            drw["nomtab"] = nomt;

                                            decimal pt = 0;
                                            pt = fn_kilosnetos_op(hisprod, var_chr_ordp_folio);

                                            drw["parcial"] = Math.Round(((peso3 * var_dec_fodp_unidades) / pt), 2);
                                            drw["tipo1"] = t1;
                                            drw["tipo2"] = t2;

                                            tmrepp1.Rows.Add(drw);

                                        }
                                        #endregion
                                    }
                                    else//Materia Prima
                                    {


                                        foreach (DataRow rw in provrchtbl.Select("rmp_recibo = '" + var_chr_rmp_recibo + "'"))
                                        {
                                            cver = rw["rch_clave"].ToString();
                                            cvet = rw["tbl_clave"].ToString();
                                            cveprov = rw["prov_clave"].ToString();

                                            nomprov = rw["rch_nombre"].ToString();
                                            nomr = rw["tbl_nombre"].ToString();
                                            nomt = rw["prov_nombre"].ToString();

                                            bool fnd = false;
                                            foreach (DataRow rwh in histo.Select("hrp_recibo = '" + var_chr_rmp_recibo + "' AND hrp_tipo_recepcion = '" + tipo + "' AND lin_clave = '" + var_lin_clave + "' AND prod_clave = '" + var_prod_clave + "'"))
                                            {
                                                peso3 = (Convert.ToDecimal(rwh["hrp_clase1"].ToString()) / Convert.ToDecimal(rwh["hrp_num_unidades"].ToString())) * var_caj_recibo;
                                                fnd = true;
                                            }
                                            if (fnd == false)
                                            {
                                                peso3 = 0;
                                            }
                                        }
                                        if (provrchtbl.Rows.Count == 0)
                                        {
                                            if (var_chr_prod_clave == "16006APPR")
                                            {
                                                MessageBox.Show("Apio proceso canastillas no encontre recibo");
                                            }
                                            cver = "";
                                            cvet = "";
                                            nomr = "";
                                            nomt = "";
                                            cveprov = "";
                                        }
                                        if ((String.Compare(cveprov, txtClaveProveedor.Text, true) > 0 || String.Compare(cveprov, txtClaveProveedor.Text, true) == 0) && (String.Compare(cveprov, txtClaveProveedor.Text, true) < 0 || String.Compare(cveprov, txtClaveProveedor.Text, true) == 0))
                                        {
                                            DataRow drw = tmrepp1.NewRow();
                                            drw["folio"] = var_chr_ordp_folio;
                                            drw["recibo"] = var_chr_rmp_recibo;
                                            drw["linea"] = var_chr_lin_clave;
                                            drw["fecha"] = var_date_ordp_fecha;
                                            drw["producto"] = var_chr_prod_clave;
                                            drw["cantidad"] = var_dec_fodp_unidades;


                                            bool fnd = false;
                                            decimal pes = 0;
                                            foreach (DataRow rx in dtequival.Select("prod_clave = '" + var_chr_prod_clave + "'"))
                                            {
                                                var_dec_um_equivalencia = Convert.ToDecimal(rx["um_equivalencia"].ToString());
                                                pes = Convert.ToDecimal(rx["compp_peso"].ToString());
                                            }
                                            if (fnd == false)
                                            {
                                                var_dec_um_equivalencia = 1;
                                            }
                                            var_dec_peso_total = var_dec_peso_total + (pes * var_dec_um_equivalencia);
                                            //var_dec_peso_total = fn_peso_total(var_chr_lin_clave, var_chr_prod_clave);

                                            drw["nombre"] = var_chr_prod_nombre;
                                            drw["tipoc"] = var_chr_tipoc;
                                            drw["neto"] = pesocom;
                                            drw["proveedor"] = cveprov;
                                            drw["nomprov"] = nomprov;
                                            drw["nombrelin"] = nom_lin;
                                            drw["cveran"] = cver;
                                            drw["nomran"] = nomr;
                                            drw["cvetab"] = cvet;
                                            drw["nomtab"] = nomt;
                                            if (var_chr_ordp_folio == "116215")
                                            {
                                            }
                                            decimal pt = 0;
                                            pt = fn_kilosnetos_op(hisprod, var_chr_ordp_folio);
                                            if (pt != 0)
                                                drw["parcial"] = Math.Round(((peso3 * var_dec_fodp_unidades) / pt), 2);
                                            else
                                                drw["parcial"] = 0;

                                            drw["tipo1"] = t1;
                                            drw["tipo2"] = t2;

                                            tmrepp1.Rows.Add(drw);
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
                reader1.Close();
                reader1.Dispose();
                thisConnection.Close();

                DataView dw = tmrepp1.DefaultView;
                dw.Sort = "producto";
                tmrepp1 = dw.ToTable();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
                return;
            }

            DateTime fecha1;
            DateTime fecha2;
            string mpro = "";

            fecha1 = Convert.ToDateTime(Convert.ToDateTime(dtpFecha1.Text).ToString("yyyy-MM-dd"));
            fecha2 = Convert.ToDateTime(Convert.ToDateTime(dtpFecha2.Text).ToString("yyyy-MM-dd"));

            mpro = txtClaveProveedor.Text;

            //DataTable tmrepp2 = new DataTable();
            ////DataTable tmrepp3 = new DataTable();
            ////DataTable tmrepp4 = new DataTable();

            //tmrepp2 = tmrepp1;
            ////tmrepp3 = tmrepp1;
            ////tmrepp4 = tmrepp1;

            //DataView dw = tmrepp2.DefaultView;
            //dw.Sort = "producto";
            //tmrepp2 = dw.ToTable();

            //DataView dw2 = tmrepp3.DefaultView;
            //dw2.Sort = "producto";
            //tmrepp3 = dw2.ToTable();

            //DataView dw3 = tmrepp3.DefaultView;
            //dw3.Sort = "producto";
            //tmrepp1 = dw3.ToTable();

            string pdcto = "";


            DataRow[] result = tmrepp1.Select("proveedor = '" + txtClaveProveedor.Text + "'");
            DataRow dr;
            foreach (DataRow row in result)
            {
                pdcto = row[4].ToString();
                if (tmrepor1a.Rows.Count == 0)
                {
                    dr = tmrepor1a.NewRow();
                    dr["linea"] = row[3].ToString();
                    dr["producto"] = row[4].ToString();
                    dr["nombrelin"] = row[8].ToString();
                    dr["nombre"] = row[7].ToString();
                    dr["neto"] = row[12].ToString();
                    dr["cajas"] = row[15].ToString();
                    tmrepor1a.Rows.Add(dr);
                }
                else
                {
                    bool fnd2 = false;
                    for (int i = 0; i < tmrepor1a.Rows.Count; i++)
                    {
                        if (pdcto == tmrepor1a.Rows[i][1].ToString())
                        {
                            tmrepor1a.Rows[i][5] = Convert.ToDecimal(tmrepor1a.Rows[i][5].ToString()) + Convert.ToDecimal(row[15].ToString());
                            fnd2 = true;
                            break;
                        }
                    }
                    if (fnd2 == false)
                    {
                        dr = tmrepor1a.NewRow();
                        dr["linea"] = row[3].ToString();
                        dr["producto"] = row[4].ToString();
                        dr["nombrelin"] = row[8].ToString();
                        dr["nombre"] = row[7].ToString();
                        dr["neto"] = row[12].ToString();
                        dr["cajas"] = row[15].ToString();
                        tmrepor1a.Rows.Add(dr);
                    }
                }


            }




            DataRow drt;
            foreach (DataRow drw in tmrepor1a.Rows)
            {
                drt = tprod.NewRow();
                drt["pro_clave"] = drw[1].ToString();
                drt["pro_nombre"] = drw[3].ToString();
                drt["lin_cve"] = drw[0].ToString();
                drt["lin_nom"] = drw[2].ToString();
                drt["unidades"] = Convert.ToDecimal(drw[5].ToString());
                drt["uninac"] = 0;
                drt["uniexp"] = 0;
                drt["liquidado"] = false;
                drt["num_liq"] = "";
                drt["neto"] = Convert.ToDecimal(drw[5].ToString());
                drt["nal"] = "0";
                drt["ex"] = "0";
                drt["LN"] = drw["linea"];

                //NUEVO CODIGO PARA CARGAR LAS CANTIDADES DE NACIONAL Y EXPORTACION
                foreach (DataRow rs in totalventas.Rows)
                {
                    bool fnd = false;
                    if (drw[1].ToString() == rs["prod_clave"].ToString())
                    {
                        fnd = true;
                        decimal nal = 0;
                        decimal exp = 0;
                        decimal rec = 0;
                        rec = Convert.ToDecimal(drt["unidades"].ToString());
                        nal = rec * Convert.ToDecimal(rs["nalpor"].ToString());
                        exp = rec * Convert.ToDecimal(rs["exppor"].ToString());
                        drt["uninac"] = (nal == 0) ? "0.00" : Math.Round(nal, 0).ToString("###,###,##0.00");//nal.ToString("###,###,###");
                        drt["uniexp"] = (exp == 0) ? "0.00" : Math.Round(exp, 0).ToString("###,###,##0.00"); //exp.ToString("###,###,###");

                        decimal total_por = rec;
                        decimal cien_por = Convert.ToDecimal("100");
                        decimal cajas_por = Convert.ToDecimal(rs["total"].ToString());
                        decimal por_total = 0;
                        if (total_por != 0)
                            por_total = ((cajas_por * cien_por) / total_por);
                        else
                            por_total = 0;

                        drw["POR"] = Math.Round(por_total, 2).ToString() + "%";
                        drw["N"] = Math.Round(Convert.ToDecimal(rs["nacional"].ToString()), 0).ToString();
                        drw["E"] = Math.Round(Convert.ToDecimal(rs["exportacion"].ToString()), 0).ToString();

                        drt["POR"] = Math.Round(por_total, 2).ToString() + "%";
                        drt["N"] = Math.Round(Convert.ToDecimal(rs["nacional"].ToString()), 0).ToString();
                        drt["E"] = Math.Round(Convert.ToDecimal(rs["exportacion"].ToString()), 0).ToString(); ;

                    }
                    if (fnd == false)
                    {
                        drt["POR"] = "0%";
                        drt["N"] = "0";
                        drt["E"] = "0";
                    }
                }

                //FIN NUEVO CODIGO

                try
                {
                    DateTime mfec;
                    mfec = Convert.ToDateTime(dtpFecha1.Text);
                    thisConnection.Open();

                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT uni_nac, uni_exp, liq_folio FROM tb_mstr_liquidacion WHERE (liq_fecha2 BETWEEN '" + fecha1.ToShortDateString() + "' AND '" + fecha2.ToShortDateString() + "')" +
                        " AND liq_lincve = '" + drw[0].ToString() + "' AND liq_prodcve = '" + drw[1].ToString() + "' AND liq_provcve = '" + mpro + "' AND status = 'A' AND tipo = 'PRO'";
                    reader1 = cmnd1.ExecuteReader();
                    while (reader1.Read())
                    {
                        drt["nal"] = reader1.GetDecimal(0);
                        drt["ex"] = reader1.GetDecimal(1);
                        drt["liquidado"] = (reader1.GetDecimal(0) > 0 || reader1.GetDecimal(1) > 0) ? true : false;
                        drt["num_liq"] = (reader1.GetDecimal(0) > 0 || reader1.GetDecimal(1) > 0) ? reader1.GetValue(2).ToString().Trim() : "";
                    }
                    reader1.Close();
                    reader1.Dispose();
                    tprod.Rows.Add(drt);

                    thisConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (thisConnection.State == ConnectionState.Open)
                        thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
                }
            }

            //for (int i = 0; i < tprod.Rows.Count; i++)
            foreach (DataRow rw in tprod.Select("lin_cve >= '" + ln1 + "' AND lin_cve < = '" + ln2 + "'"))
            {
                //dtgLiquidacion.Rows.Add(tprod.Rows[i][0].ToString(), tprod.Rows[i][1].ToString(), tprod.Rows[i][3].ToString(), tprod.Rows[i][9].ToString(), tprod.Rows[i][5].ToString(), tprod.Rows[i][6].ToString(), tprod.Rows[i][8].ToString(), Convert.ToBoolean(tprod.Rows[i][7].ToString()), tprod.Rows[i][10].ToString(), tprod.Rows[i][11].ToString(), tprod.Rows[i][12].ToString(), tprod.Rows[i][13].ToString(), tprod.Rows[i][14].ToString(), tprod.Rows[i][15].ToString(), tprod.Rows[i][16].ToString(), tprod.Rows[i][17].ToString(), Convert.ToDecimal(tprod.Rows[i][18].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(tprod.Rows[i][19].ToString()).ToString("###,###,##0.000"));
                dtgLiquidacion.Rows.Add(rw["pro_clave"].ToString(), rw["pro_nombre"].ToString(), rw["lin_nom"].ToString(), Convert.ToDecimal(rw["unidades"].ToString()).ToString("###,###,##0.00"),
                    Convert.ToDecimal(rw["uninac"].ToString()).ToString("###,###,##0.00"), Convert.ToDecimal(rw["uniexp"].ToString()).ToString("###,###,##0.00"), rw["num_liq"].ToString(),
                    Convert.ToBoolean(rw["liquidado"].ToString()), Convert.ToDecimal(rw["nal"].ToString()).ToString("###,###,##0.00"), Convert.ToDecimal(rw["ex"].ToString()).ToString("###,###,##0.00"),
                    rw["POR"].ToString(), rw["N"].ToString(), rw["E"].ToString(), rw["lin_cve"].ToString());
            }

            colorear();

            string filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Consulta de liquidaciones de materia prima");
                sw.Close();
            }
            btnGenera.Enabled = true;
            label14.Visible = false;

        }

        private void DDLProveedor_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLProveedor.SelectedIndex > -1)
            {
                txtClaveProveedor.Text = proveedor.Rows[DDLProveedor.SelectedIndex]["prov_clave"].ToString();
                dtpFecha1.Focus();
            }
        }

        private void txtClaveProveedor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                int i = 0;
                foreach (DataRow rw in proveedor.Rows)
                {
                    if (rw["prov_clave"].ToString() == txtClaveProveedor.Text)
                    {
                        DDLProveedor.SelectedIndex = i;
                        dtpFecha1.Focus();
                        break;
                    }
                    i++;
                }
            }
        }

        private bool validarfecha(string fecha)
        {
            try
            {
                string f = Convert.ToDateTime(fecha).ToShortDateString();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato de la fecha no es válido", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private void dtpFecha1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (validarfecha(dtpFecha1.Text) == false)
                {
                    dtpFecha1.Text = "";
                    return;
                }
                else
                {
                    dtpFecha2.Focus();
                }
            }
        }

        private void dtpFecha2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (validarfecha(dtpFecha2.Text) == false)
                {
                    dtpFecha2.Text = "";
                    return;
                }
                else
                {
                    txtLinea1.Focus();
                }
            }
        }

        private void txtLinea1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                int i = 0;
                foreach (DataRow rw in lineas.Rows)
                {
                    if (rw["lin_clave"].ToString() == txtLinea1.Text)
                    {
                        DDLLinea1.SelectedIndex = i;
                        txtLinea2.Focus();
                        foreach (DataRow rz in productos.Select("lin_clave = '" + txtLinea1.Text + "'"))
                        {
                            DDLEmpaques1.Items.Add(rz["prod_nombre"].ToString());
                        }
                        break;
                    }
                    i++;
                }
            }
        }

        private void txtLinea2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                int i = 0;
                foreach (DataRow rw in lineas.Rows)
                {
                    if (rw["lin_clave"].ToString() == txtLinea2.Text)
                    {
                        DDLLinea2.SelectedIndex = i;
                        txtEmpaque1.Focus();
                        foreach (DataRow rz in productos.Select("lin_clave = '" + txtLinea2.Text + "'"))
                        {
                            DDLEmpaques2.Items.Add(rz["prod_nombre"].ToString());
                        }
                        break;
                    }
                    i++;
                }
            }
        }

        private void pbxCalendarUlEn_Click(object sender, EventArgs e)
        {
            calendario dlg = new calendario();
            //dlg.Location = new Point(pbxCalendarUlEn.Location.X + 313, pbxCalendarUlEn.Location.Y + 270);

            DialogResult dr = dlg.ShowDialog();
            if (dr == DialogResult.OK)
            {
                string fechaUlEn = liquidacion.calendario.SharedData.Polino.fecha;
                dtpFecha1.Text = Convert.ToDateTime(fechaUlEn).ToString("dd/MM/yyyy");
                dtpFecha2.Focus();
            }
        }

        private void pbxFecha2_Click(object sender, EventArgs e)
        {
            calendario dlg = new calendario();
            //dlg.Location = new Point(pbxCalendarUlEn.Location.X + 553, pbxCalendarUlEn.Location.Y + 270);

            DialogResult dr = dlg.ShowDialog();
            if (dr == DialogResult.OK)
            {
                string fechaUlEn = liquidacion.calendario.SharedData.Polino.fecha;
                dtpFecha2.Text = Convert.ToDateTime(fechaUlEn).ToString("dd/MM/yyyy");
                txtLinea1.Focus();
            }
        }

        private void btnMinimizar_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;

        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            try
            {
                thisConnection.Open();
                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema = 'SISEMP' ORDER BY inicio_sesion desc";
                reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader2.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader2.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader2.Close();

                cmnd2 = thisConnection.CreateCommand();
                cmnd2.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema = 'SISEMP'";
                reader2 = cmnd2.ExecuteReader();
                reader2.Close();
                thisConnection.Close();

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Salida 4.1 Liquidaciones");
                    sw.Close();
                }

                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", ex.ToString(), "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
            }
        }

        private void DDLLinea1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLProveedor.SelectedIndex > -1)
            {
                txtLinea1.Text = lineas.Rows[DDLLinea1.SelectedIndex]["lin_clave"].ToString();
                KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                txtLinea1_KeyPress(sender, llave);
                txtLinea2.Focus();
            }
        }

        private void DDLLinea2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLLinea2.SelectedIndex > -1)
            {
                txtLinea2.Text = lineas.Rows[DDLLinea2.SelectedIndex]["lin_clave"].ToString();
                KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                txtLinea2_KeyPress(sender, llave);
                txtEmpaque1.Focus();
            }
        }

        private void DDLEmpaques1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLProveedor.SelectedIndex > -1)
            {
                txtEmpaque1.Text = productos.Rows[DDLEmpaques1.SelectedIndex]["prod_clave"].ToString();
                KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                txtEmpaque1_KeyPress(sender, llave);
                txtEmpaque2.Focus();
            }
        }

        private void DDLEmpaques2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (DDLProveedor.SelectedIndex > -1)
            {
                txtEmpaque2.Text = productos.Rows[DDLEmpaques2.SelectedIndex]["prod_clave"].ToString();
            }
        }

        private void txtEmpaque1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                int i = 0;
                foreach (DataRow rw in productos.Rows)
                {
                    if (rw["prod_clave"].ToString() == txtEmpaque1.Text)
                    {
                        DDLEmpaques1.SelectedIndex = i;
                        txtEmpaque2.Focus();
                        break;
                    }
                    i++;
                }
            }
        }

        private void txtEmpaque2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                int i = 0;
                foreach (DataRow rw in productos.Rows)
                {
                    if (rw["prod_clave"].ToString() == txtEmpaque2.Text)
                    {
                        DDLEmpaques2.SelectedIndex = i;
                        break;
                    }
                    i++;
                }
            }
        }

        public string nombreproducto(string clavep)
        {
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prod_nombre FROM tb_cat_producto WHERE prod_clave = '" + clavep + "' ORDER BY prod_clave";
            reader1 = cmnd1.ExecuteReader();
            string nom = "";
            if (reader1.HasRows)
            {
                reader1.Read();
                nom = reader1.GetValue(0).ToString().Trim();
            }
            thisConnection.Close();
            return nom;
        }

        public string nombrelinea(string clavep)
        {
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT lin_nombre FROM tb_cat_linea WHERE lin_clave = '" + clavep + "' ORDER BY lin_clave";
            reader1 = cmnd1.ExecuteReader();
            string nom = "";
            if (reader1.HasRows)
            {
                reader1.Read();
                nom = reader1.GetValue(0).ToString().Trim();
            }
            thisConnection.Close();
            return nom;
        }

        public void cargadatosliquidaciones()
        {
            datosliq.Clear();
            thisConnection.Open();
            string fecha = DateTime.Now.AddMonths(-5).ToShortDateString();
            //string fecha = DateTime.Now.AddMonths(-5).ToString("yyyy-MM-dd");
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT liq_folio, uni_nac, uni_exp FROM tb_mstr_liquidacion WHERE liq_fecha >= '" + Convert.ToDateTime(dtpFecha1.Text).ToShortDateString() + "' ORDER BY liq_folio";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            while (reader1.Read())
            {
                rw = datosliq.NewRow();
                rw["liq_folio"] = reader1.GetValue(0).ToString().Trim();
                rw["uni_nac"] = reader1.GetValue(1).ToString().Trim();
                rw["uni_exp"] = reader1.GetValue(2).ToString().Trim();
                datosliq.Rows.Add(rw);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
        }

        private void colorear()
        {
            foreach (DataGridViewRow rw in dtgLiquidacion.Rows)
            {
                decimal cant = Convert.ToDecimal(rw.Cells[3].Value.ToString());
                decimal nal = Convert.ToDecimal(rw.Cells[4].Value.ToString());
                decimal exp = Convert.ToDecimal(rw.Cells[5].Value.ToString());
                decimal nal_liq = Convert.ToDecimal(rw.Cells[8].Value.ToString());
                decimal exp_liq = Convert.ToDecimal(rw.Cells[9].Value.ToString());

                string liqnum = rw.Cells[6].Value.ToString();

                //colorear grid cantidades
                if (nal > 0 && exp > 0)
                {
                    if (nal == nal_liq)
                        rw.Cells[4].Style.BackColor = Color.Green;
                    else
                        rw.Cells[4].Style.BackColor = Color.Red;
                    if (exp == exp_liq)
                        rw.Cells[5].Style.BackColor = Color.Green;
                    else
                        rw.Cells[5].Style.BackColor = Color.Red;
                }
                if (nal > 0 && exp == 0)
                {
                    if (nal == nal_liq)
                        rw.Cells[4].Style.BackColor = Color.Green;
                    else
                        rw.Cells[4].Style.BackColor = Color.Red;
                }
                if (nal == 0 && exp > 0)
                {
                    if (exp == exp_liq)
                        rw.Cells[5].Style.BackColor = Color.Green;
                    else
                        rw.Cells[5].Style.BackColor = Color.Red;
                }

            }
        }

        private void dtgLiquidacion_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (nanana == false)
            {
                MessageBox.Show("Recuerde que no se ha realizado el calculo de los fletes", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //return;
            }
            string tl = "";
            if (DDLTipo.SelectedIndex == 0)
                tl = "PTC";
            if (DDLTipo.SelectedIndex == 1)
                tl = "PRO";
            //NACIONALES
            #region nacional
            if (e.ColumnIndex == 4)
            {
                lblRenglon.Text = e.RowIndex.ToString();
                //NUEVA
                if (lblBloqueo.Text == "1")
                {
                    if (txtClaveProveedor.Text != "01" && txtClaveProveedor.Text != "03" && txtClaveProveedor.Text != "1328")
                    {
                        if (dtgLiquidacion.CurrentRow.Cells["ln"].Value.ToString() != "12" && dtgLiquidacion.CurrentRow.Cells["ln"].Value.ToString() != "19")
                        {
                            string por = dtgLiquidacion.CurrentRow.Cells["por"].Value.ToString().TrimEnd('%');
                            decimal porcentaje_venta = Convert.ToDecimal(por);
                            string N = dtgLiquidacion.CurrentRow.Cells["n"].Value.ToString();
                            string E = dtgLiquidacion.CurrentRow.Cells["e"].Value.ToString();
                            if (porcentaje_venta < 65)
                            {
                                MessageBox.Show("La venta del producto no supera el 65%, no es válido para realizar la liquidación.\nCajas Ventas Nacional: " + N + "\nCajas Ventas Exportación: " + E, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }
                }



                if (dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString() == "")
                {
                    tipo = "nueva";
                    cveprov = txtClaveProveedor.Text;
                    nomprov = DDLProveedor.SelectedItem.ToString();
                    cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nac"].Value.ToString()).ToString();
                    num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();
                    cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                    nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                    cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                    nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                    f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                    f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                    procedencia = "NACIONAL";
                    neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                    if (DDLTipo.SelectedIndex == 1)
                    {
                        dtrecibos.Clear();

                        DataView dw = tmrepp1.DefaultView;
                        dw.RowFilter = "producto = '" + cveprod + "'";
                        dtrecibos = dw.ToTable();
                    }

                    if (MessageBox.Show("Desea introducir el número de cajas", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        cajas dlg2 = new cajas();
                        dlg2.ShowDialog();
                        if (dlg2.DialogResult == DialogResult.Yes)
                        {
                            cantidad = liquidacion.cajas.SharedData.Polino.cantidad;
                            preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                            dlg.ShowDialog();

                            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                            {
                                if (DDLTipo.SelectedIndex == 0)
                                {
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                                else
                                {
                                    btnGenera_Click(null, null);
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                            }
                        }
                        else
                        {
                            return;
                        }

                    }
                    else
                    {

                        cantidad = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nac"].Value.ToString()).ToString();
                        preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                        dlg.ShowDialog();

                        if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                        {
                            //btnGenera_Click(null, null);
                            //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            if (DDLTipo.SelectedIndex == 0)
                            {
                                btnGenera_Click(null, null);
                                dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            }
                            else
                            {
                                btnGenera_Click(null, null);
                                btnGenera_Click(null, null);
                                dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            }
                        }

                    }
                    return;
                }

                //CONSULTA
                if (dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nal"].Value.ToString()) == 0)
                    {
                        tipo = "nuevotipo";
                        cveprov = txtClaveProveedor.Text;
                        nomprov = DDLProveedor.SelectedItem.ToString();
                        cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nac"].Value.ToString()).ToString();
                        num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();

                        cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                        nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                        cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                        nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                        f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                        f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                        procedencia = "NACIONAL";
                        neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                        if (DDLTipo.SelectedIndex == 1)
                        {
                            dtrecibos.Clear();

                            DataView dw = tmrepp1.DefaultView;
                            dw.RowFilter = "producto = '" + cveprod + "'";
                            dtrecibos = dw.ToTable();
                        }

                        if (MessageBox.Show("Desea introducir el numero de cajas", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            cajas dlg2 = new cajas();
                            dlg2.ShowDialog();
                            if (dlg2.DialogResult == DialogResult.Yes)
                            {
                                cantidad = liquidacion.cajas.SharedData.Polino.cantidad;
                                preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                                dlg.ShowDialog();

                                if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                                {
                                    //btnGenera_Click(null, null);
                                    //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    if (DDLTipo.SelectedIndex == 0)
                                    {
                                        btnGenera_Click(null, null);
                                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    }
                                    else
                                    {
                                        btnGenera_Click(null, null);
                                        btnGenera_Click(null, null);
                                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    }
                                }
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            if (DDLTipo.SelectedIndex == 1)
                            {
                                dtrecibos.Clear();

                                DataView dw = tmrepp1.DefaultView;
                                dw.RowFilter = "producto = '" + cveprod + "'";
                                dtrecibos = dw.ToTable();
                            }

                            cantidad = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nac"].Value.ToString()).ToString();
                            preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                            dlg.ShowDialog();

                            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                            {
                                //btnGenera_Click(null, null);
                                //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                if (DDLTipo.SelectedIndex == 0)
                                {
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                                else
                                {
                                    btnGenera_Click(null, null);
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                            }
                            //preliminar dlg = new preliminar(cantidad, tipo, cvelin, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia);
                            //dlg.ShowDialog();
                        }
                    }
                    else
                    {
                        tipo = "consulta";
                        cveprov = txtClaveProveedor.Text;
                        nomprov = DDLProveedor.SelectedItem.ToString();
                        cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["nac"].Value.ToString()).ToString();
                        num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();

                        cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                        nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                        cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                        nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                        f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                        f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                        procedencia = "NACIONAL";
                        neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                        cantidad = dtgLiquidacion.Rows[e.RowIndex].Cells["nal"].Value.ToString();

                        //if (DDLTipo.SelectedIndex == 1)
                        //{
                        //    dtrecibos.Clear();

                        //    DataView dw = tmrepp1.DefaultView;
                        //    dw.RowFilter = "producto = '" + cveprod + "'";
                        //    dtrecibos = dw.ToTable();
                        //}

                        preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                        dlg.ShowDialog();

                        btnGenera_Click(null, null);
                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                    }

                }
            }
            #endregion

            //EXPORTACION
            #region exportacion
            if (e.ColumnIndex == 5)
            {
                lblRenglon.Text = e.RowIndex.ToString();

                if (lblBloqueo.Text == "1")
                {
                    if (txtClaveProveedor.Text != "01" && txtClaveProveedor.Text != "03" && txtClaveProveedor.Text != "1328")
                    {
                        if (dtgLiquidacion.CurrentRow.Cells["ln"].Value.ToString() != "12" && dtgLiquidacion.CurrentRow.Cells["ln"].Value.ToString() != "19")
                        {
                            string por = dtgLiquidacion.CurrentRow.Cells["por"].Value.ToString().TrimEnd('%');
                            decimal porcentaje_venta = Convert.ToDecimal(por);
                            string N = dtgLiquidacion.CurrentRow.Cells["n"].Value.ToString();
                            string E = dtgLiquidacion.CurrentRow.Cells["e"].Value.ToString();
                            if (porcentaje_venta < 65)
                            {
                                MessageBox.Show("La venta del producto no supera el 65%, no es válido para realizar la liquidación.\nCajas Ventas Nacional: " + N + "\nCajas Ventas Exportación: " + E, "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }
                }


                //NUEVA
                if (dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString() == "")
                {
                    tipo = "nueva";
                    cveprov = txtClaveProveedor.Text;
                    nomprov = DDLProveedor.SelectedItem.ToString();
                    cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["exp"].Value.ToString()).ToString();
                    num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();
                    cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                    nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                    cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                    nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                    f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                    f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                    procedencia = "EXPORTACION";
                    neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                    if (DDLTipo.SelectedIndex == 1)
                    {
                        dtrecibos.Clear();

                        DataView dw = tmrepp1.DefaultView;
                        dw.RowFilter = "producto = '" + cveprod + "'";
                        dtrecibos = dw.ToTable();
                    }

                    if (MessageBox.Show("Desea introducir el número de cajas", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        cajas dlg2 = new cajas();
                        dlg2.ShowDialog();
                        if (dlg2.DialogResult == DialogResult.Yes)
                        {
                            cantidad = liquidacion.cajas.SharedData.Polino.cantidad;
                            preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                            dlg.ShowDialog();

                            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                            {
                                //btnGenera_Click(null, null);
                                //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                if (DDLTipo.SelectedIndex == 0)
                                {
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                                else
                                {
                                    btnGenera_Click(null, null);
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                            }
                        }
                        else
                        {
                            return;
                        }

                    }
                    else
                    {
                        cantidad = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["exp"].Value.ToString()).ToString();
                        preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                        dlg.ShowDialog();

                        if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                        {
                            //btnGenera_Click(null, null);
                            //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            if (DDLTipo.SelectedIndex == 0)
                            {
                                btnGenera_Click(null, null);
                                dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            }
                            else
                            {
                                btnGenera_Click(null, null);
                                btnGenera_Click(null, null);
                                dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                            }
                        }
                    }
                    return;
                }

                //CONSULTA
                if (dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["ex"].Value.ToString()) == 0)
                    {
                        tipo = "nuevotipo";
                        cveprov = txtClaveProveedor.Text;
                        nomprov = DDLProveedor.SelectedItem.ToString();
                        cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["exp"].Value.ToString()).ToString();
                        num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();

                        cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                        nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                        cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                        nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                        f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                        f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                        procedencia = "EXPORTACION";
                        neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                        if (DDLTipo.SelectedIndex == 1)
                        {
                            dtrecibos.Clear();

                            DataView dw = tmrepp1.DefaultView;
                            dw.RowFilter = "producto = '" + cveprod + "'";
                            dtrecibos = dw.ToTable();
                        }

                        if (MessageBox.Show("Desea introducir el numero de cajas", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            cajas dlg2 = new cajas();
                            dlg2.ShowDialog();
                            if (dlg2.DialogResult == DialogResult.Yes)
                            {
                                cantidad = liquidacion.cajas.SharedData.Polino.cantidad;
                                preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                                dlg.ShowDialog();

                                if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                                {
                                    //btnGenera_Click(null, null);
                                    //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    if (DDLTipo.SelectedIndex == 0)
                                    {
                                        btnGenera_Click(null, null);
                                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    }
                                    else
                                    {
                                        btnGenera_Click(null, null);
                                        btnGenera_Click(null, null);
                                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                    }
                                }
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            cantidad = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["exp"].Value.ToString()).ToString();
                            preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                            dlg.ShowDialog();

                            if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                            {
                                //btnGenera_Click(null, null);
                                //dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                if (DDLTipo.SelectedIndex == 0)
                                {
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                                else
                                {
                                    btnGenera_Click(null, null);
                                    btnGenera_Click(null, null);
                                    dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                                }
                            }
                            //preliminar dlg = new preliminar(cantidad, tipo, cvelin, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia);
                            //dlg.ShowDialog();
                        }
                    }
                    else
                    {
                        tipo = "consulta";
                        cveprov = txtClaveProveedor.Text;
                        nomprov = DDLProveedor.SelectedItem.ToString();
                        cant_orig = Convert.ToDecimal(dtgLiquidacion.Rows[e.RowIndex].Cells["exp"].Value.ToString()).ToString();
                        num_liq = dtgLiquidacion.Rows[e.RowIndex].Cells["nliq"].Value.ToString();

                        cveprod = dtgLiquidacion.Rows[e.RowIndex].Cells["producto"].Value.ToString();
                        nomprod = dtgLiquidacion.Rows[e.RowIndex].Cells["nombre"].Value.ToString();
                        cvelin = tprod.Rows[e.RowIndex]["lin_cve"].ToString();
                        nomlin = tprod.Rows[e.RowIndex]["lin_nom"].ToString();
                        f1 = Convert.ToDateTime(dtpFecha1.Text).ToShortDateString();
                        f2 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
                        procedencia = "EXPORTACION";
                        neto_prod = tprod.Rows[e.RowIndex]["neto"].ToString();

                        cantidad = dtgLiquidacion.Rows[e.RowIndex].Cells["ex"].Value.ToString();

                        if (DDLTipo.SelectedIndex == 1)
                        {
                            dtrecibos.Clear();

                            DataView dw = tmrepp1.DefaultView;
                            dw.RowFilter = "producto = '" + cveprod + "'";
                            dtrecibos = dw.ToTable();
                        }

                        preliminar dlg = new preliminar(cantidad, tipo, num_liq, cant_orig, cveprod, nomprod, cvelin, nomlin, f1, f2, cveprov, nomprov, procedencia, neto_prod, tl, dtrecibos);
                        dlg.ShowDialog();

                        btnGenera_Click(null, null);
                        btnGenera_Click(null, null);
                        dtgLiquidacion.FirstDisplayedScrollingRowIndex = Convert.ToInt32(lblRenglon.Text);
                    }

                }
            }
            #endregion
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

        public decimal fn_kilosnetos_op(DataTable tb, string cverecibo)
        {
            pesotot = 0;
            decimal num = 0;
            foreach (DataRow rw in tb.Select("ordp_folio = '" + cverecibo + "'"))
            {
                num = Convert.ToDecimal(Convert.ToDecimal(rw["hrp_clase1"].ToString()) / Convert.ToDecimal(rw["hrp_num_unidades"].ToString())) * Convert.ToDecimal(rw["podp_cantidad"].ToString());
                pesotot = pesotot + num;
            }
            //cmnd2 = thisConnection.CreateCommand();
            //cmnd2.CommandText = "SELECT P.rmp_recibo, P.rmp_tipo, P.lin_clave, P.prod_clave, SUM((H.hrp_clase1 / H.hrp_num_unidades) * P.podp_cantidad) AS pesotot" +
            //    " FROM tb_det_prod_odp P, tb_hist_recepcion H" +
            //    " WHERE P.rmp_recibo = H.hrp_recibo AND P.lin_clave = H.lin_clave AND P.rmp_tipo = H.hrp_tipo_recepcion AND P.lin_clave = H.lin_clave" +
            //    " AND P.prod_clave = H.prod_clave AND P.ordp_folio = '" + cverecibo + "'" +
            //    " group by P.rmp_recibo, P.rmp_tipo, P.prod_clave, P.lin_clave order by rmp_recibo, P.rmp_tipo, P.prod_clave, P.lin_clave";
            //reader2 = cmnd2.ExecuteReader();
            //if (reader2.HasRows)
            //{
            //    while (reader2.Read())
            //    {
            //        //pesotot = pesotot + ((reader7.GetDecimal(5) / reader7.GetDecimal(6)) * reader7.GetDecimal(0));
            //        pesotot = pesotot + reader2.GetDecimal(4);
            //    }
            //}
            //reader2.Close();
            //reader2.Dispose();
            return pesotot;
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            DDLProveedor.Enabled = true;
            txtClaveProveedor.ReadOnly = false;
            tprod.Clear();
            dtgLiquidacion.Rows.Clear();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        public string ultima_fecha_calculada()
        {
            string fch = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            if (tipo_reporte == "M")
            {
                fch = "Fecha ultima de calculo mensual de fletes ";
                cmnd1.CommandText = "SELECT TOP 1 fecha FROM tb_registro_movimientos WHERE tipo_mov = 'F' AND folio = 'MES' ORDER BY fecha DESC ";
            }
            else
            {
                fch = "Fecha ultima de calculo semanal de fletes ";
                cmnd1.CommandText = "SELECT TOP 1 fecha FROM tb_registro_movimientos WHERE tipo_mov = 'F' AND folio = 'SEM' ORDER BY fecha DESC ";
            }

            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                fch += Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToString("dddd, dd MMMM yyyy HH:mm");
                fech_ulti = Convert.ToDateTime(reader1["fecha"].ToString().Trim()).ToShortDateString();
            }
            else
            {
                fch += "No existe registro";
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            return fch;
        }

        private void btnImpuesto_Click(object sender, EventArgs e)
        {
            Notas_Credito_Cargo dlg = new Notas_Credito_Cargo(tprod, dtpFecha1.Text, dtpFecha2.Text, txtClaveProveedor.Text, dtLineas);
            dlg.Show();
        }


    }
}
