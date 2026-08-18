using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Globalization;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

namespace liquidacion
{
    public partial class preliminar : Form
    {
        //SqlConnection thisConnection = new SqlConnection("Data Source=GABIRA1\\SQL2005;Initial Catalog=GAB_Irapuato;Connect Timeout=130;User ID=sa; MultipleActiveResultSets=True");
        SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
        SqlDataReader reader1, reader2, reader3;
        SqlCommand cmnd3;
        SqlCommand cmnd2;
        SqlCommand cmnd1;

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
        string procedencia = "";
        string neto_prod = "";

        string filelog = "";

        string tl = "";

        DataTable tpp = new DataTable();
        DataTable tppexp = new DataTable();
        DataTable detliq = new DataTable();
        DataTable tcon = new DataTable();
        DataTable tppe = new DataTable();
        DataTable ecto = new DataTable();
        DataTable tablanc = new DataTable();
        DataTable tbrecibos = new DataTable();

        List<string> clavedesc = new List<string>();
        List<string> nombredesc = new List<string>();

        DataTable dtrecibos = new DataTable();

        DataTable recibs = new DataTable();

        string ultimo_folio = "";

        DataTable dtPrestamos = new DataTable();

        string nal_dlls = "";



        public preliminar(string can, string tip, string nli, string ori, string cprod, string nprod, string clin, string nlin, string fe1, string fe2, string cprov, string nprov, string proc, string neto, string ttl, DataTable rec)
        {
            InitializeComponent();

            dtgConceptos.EnableHeadersVisualStyles = false;
            dtgConceptos.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;

            string ruta = @"C:\SisGabWeb\fondo_formularios.jpg";
            this.BackgroundImage = System.Drawing.Bitmap.FromFile(ruta);

            filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Acceso a liquidación preliminar");
                sw.Close();
            }

            //----------29/11/2017----------//
            //VERIFICAR TIPO DE CAMBIO AL DIA
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            string dia = DateTime.Now.Day.ToString();
            string mes = DateTime.Now.Month.ToString();
            string anio = DateTime.Now.Year.ToString();
            cmnd1.CommandText = "SELECT valor, fecha FROM tb_cat_tipocambio WHERE dia = '" + dia + "' AND mes = '" + mes + "' AND año = '" + anio + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                lblTipoCambio.Text = reader1["valor"].ToString().Trim();
            }
            else
            {
                MessageBox.Show("No se ha dado de alta el tipo de cambio del día actual, favor de verificarlo con el encargado", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                thisConnection.Close();
                return;
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            //FIN VERIFICAR TIPO DE CAMBIO AL DIA
            //----------29/11/2017----------//

            tpp.Columns.Add("fcn_folio", typeof(string));
            tpp.Columns.Add("fcn_tipo", typeof(string));
            tpp.Columns.Add("fcn_estatus", typeof(string));
            tpp.Columns.Add("unidades", typeof(decimal));
            tpp.Columns.Add("precio_mn", typeof(decimal));
            tpp.Columns.Add("lin_clave", typeof(string));
            tpp.Columns.Add("prod_clave", typeof(string));

            tppexp.Columns.Add("fcn_folio", typeof(string));
            tppexp.Columns.Add("fcn_tipo", typeof(string));
            tppexp.Columns.Add("fcn_estatus", typeof(string));
            tppexp.Columns.Add("unidades", typeof(decimal));
            tppexp.Columns.Add("precio_usd", typeof(decimal));
            tppexp.Columns.Add("lin_clave", typeof(string));
            tppexp.Columns.Add("prod_clave", typeof(string));

            detliq.Columns.Add("ordp_folio", Type.GetType("System.String"));
            detliq.Columns.Add("prod_clave", Type.GetType("System.String"));
            detliq.Columns.Add("prov_clave", Type.GetType("System.String"));
            detliq.Columns.Add("liq_folio", Type.GetType("System.String"));
            detliq.Columns.Add("liq_cant", Type.GetType("System.Decimal"));
            detliq.Columns.Add("liq_fecha", Type.GetType("System.DateTime"));
            detliq.Columns.Add("liq_NoE", Type.GetType("System.String"));
            detliq.Columns.Add("estatus", Type.GetType("System.String"));

            tcon.Columns.Add("cve_con", Type.GetType("System.String"));//0
            tcon.Columns.Add("nombre_con", Type.GetType("System.String"));//1
            tcon.Columns.Add("unidades", Type.GetType("System.Decimal"));//2
            tcon.Columns.Add("precio", Type.GetType("System.Decimal"));//3
            tcon.Columns.Add("total", Type.GetType("System.Decimal"));//4
            tcon.Columns.Add("valor", Type.GetType("System.String"));//5
            tcon.Columns.Add("moni", Type.GetType("System.String"));//6
            tcon.Columns.Add("tc", Type.GetType("System.String"));//6
            tcon.Columns.Add("conse", Type.GetType("System.String"));//6
            tcon.Columns.Add("saldo", Type.GetType("System.String"));//6
            tcon.Columns.Add("val", Type.GetType("System.String"));//6
            tcon.Columns.Add("calculo", Type.GetType("System.String"));//6

            tppe.Columns.Add("emp_clave", typeof(string));
            tppe.Columns.Add("hrp_costo", typeof(decimal));

            ecto.Columns.Add("emp_clave", typeof(string));
            ecto.Columns.Add("emp_costo", typeof(decimal));

            tablanc.Columns.Add("prod_nombre", typeof(string));
            tablanc.Columns.Add("nc_folio", typeof(string));
            tablanc.Columns.Add("dnc_cantidad", typeof(string));
            tablanc.Columns.Add("dnc_precio_mn", typeof(string));
            tablanc.Columns.Add("dnc_precio_usd", typeof(string));
            tablanc.Columns.Add("clavep", typeof(string));
            tablanc.Columns.Add("dnc_tipo", typeof(string));
            tablanc.Columns.Add("lin_clave", typeof(string));
            tablanc.Columns.Add("fechap", typeof(string));

            tbrecibos.Columns.Add("hrp_recibo", typeof(string));
            tbrecibos.Columns.Add("hrp_cantidad", typeof(string));

            recibs.Columns.Add("Ordp_Folio", typeof(string));
            recibs.Columns.Add("Prod_cve", typeof(string));
            recibs.Columns.Add("Prov_cve", typeof(string));
            recibs.Columns.Add("Liq_Folio", typeof(string));

            dtPrestamos.Columns.Add("Id_Movimiento", typeof(string));
            dtPrestamos.Columns.Add("Fecha", typeof(string));
            dtPrestamos.Columns.Add("Prov_Clave", typeof(string));
            dtPrestamos.Columns.Add("Descripcion_Art", typeof(string));
            dtPrestamos.Columns.Add("Lin_Clave", typeof(string));
            dtPrestamos.Columns.Add("Total", typeof(string));
            dtPrestamos.Columns.Add("Saldo", typeof(string));
            dtPrestamos.Columns.Add("Moneda", typeof(string));
            dtPrestamos.Columns.Add("TipoCambio", typeof(string));
            dtPrestamos.Columns.Add("Factura", typeof(string));

            string t_i_p_o = "";
            if (proc == "NACIONAL")
                t_i_p_o = "PESOS";
            if (proc == "EXPORTACION")
                t_i_p_o = "DOLARES";

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            //cmnd1.CommandText = "SELECT Id_Movimiento, Fecha, Prov_Clave, Descripcion_Art, Lin_Clave, Total, Saldo FROM Tb_Prestamos_Prov WHERE Prov_Clave = '" + cprov + "'" +
            //    " AND Lin_Clave = '" + clin + "' AND Moneda = '" + t_i_p_o + "' AND Saldo < Total ORDER BY Fecha ASC";
            //cmnd1.CommandText = "SELECT Id_Movimiento, Fecha, Descripcion_Art, Id_Clave_Desc, Cantidad, Saldo FROM Tb_Prestamos_Prov WHERE Prov_Clave = '" + cprov + "'" +
            //    " AND Moneda = '" + t_i_p_o + "' AND Saldo < Total AND estatus <> 'C' ORDER BY Fecha ASC";
            cmnd1.CommandText = "SELECT Id_Movimiento, Fecha, Descripcion_Art, Id_Clave_Desc, Cantidad, Saldo, Moneda, factura FROM Tb_Prestamos_Prov WHERE Prov_Clave = '" + cprov + "'" +
                " AND Saldo < Total AND estatus <> 'C' ORDER BY Fecha ASC";
            reader1 = cmnd1.ExecuteReader();
            DataRow rr;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rr = dtPrestamos.NewRow();
                    rr["Id_Movimiento"] = reader1["Id_Movimiento"].ToString().Trim();
                    rr["Fecha"] = reader1["Fecha"].ToString().Trim();
                    rr["Prov_Clave"] = "";
                    rr["Descripcion_Art"] = reader1["Descripcion_Art"].ToString().Trim();
                    rr["Lin_Clave"] = reader1["Id_Clave_Desc"].ToString().Trim();
                    rr["Total"] = reader1["Cantidad"].ToString().Trim();
                    rr["Saldo"] = reader1["Saldo"].ToString().Trim();
                    rr["Moneda"] = reader1["Moneda"].ToString().Trim();
                    if (t_i_p_o == "PESOS")
                    {
                        if (reader1["Moneda"].ToString().Trim() == "DOLARES")
                            rr["TipoCambio"] = lblTipoCambio.Text;
                    }
                    if (t_i_p_o == "DOLARES")
                    {
                        if (reader1["Moneda"].ToString().Trim() == "PESOS")
                            rr["TipoCambio"] = txt_tipocambio.Text;
                    }
                    rr["Factura"] = reader1["factura"].ToString().Trim();
                    dtPrestamos.Rows.Add(rr);
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            cantidad = can;
            tipo = tip;
            num_liq = nli;
            cant_orig = ori;
            cveprod = cprod;
            nomprod = nprod;
            cvelin = clin;
            nomlin = nlin;
            f1 = fe1;
            f2 = fe2;
            cveprov = cprov;
            nomprov = nprov;
            procedencia = proc;
            neto_prod = neto;
            tl = ttl;

            lbl_liquidacion.Text = num_liq;
            lbl_cveprov.Text = cveprov;
            lbl_proveedor.Text = nomprov;
            lbl_cveprod.Text = cveprod;
            lbl_producto.Text = nomprod;
            lbl_fecha1.Text = f1;
            lbl_fecha2.Text = f2;

            txt_lincve.Text = cvelin;
            txt_linnom.Text = nomlin;

            txt_tipo.Text = procedencia;
            txtTipoLiq.Text = tipo;
            lblTeorico.Text = cant_orig;

            txtTL.Text = tl;

            dtrecibos = rec;

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT ordp_folio, prod_cve, prov_cve, liq_folio, liq_cant, liq_fecha, liq_noe, estatus FROM tb_det_liq_planta" +
                " WHERE prov_cve = '" + cveprov + "' AND prod_cve = '" + cveprod + "' AND (liq_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "'" +
                " AND '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                " AND estatus = 'A'";
            reader1 = cmnd1.ExecuteReader();
            DataRow drl;
            while (reader1.Read())
            {
                drl = detliq.NewRow();
                drl["ordp_folio"] = reader1.GetValue(0).ToString().Trim();
                drl["prod_clave"] = reader1.GetValue(1).ToString().Trim();
                drl["prov_clave"] = reader1.GetValue(2).ToString().Trim();
                drl["liq_folio"] = reader1.GetValue(3).ToString().Trim();
                drl["liq_cant"] = reader1.GetDecimal(4);
                drl["liq_fecha"] = reader1.GetDateTime(5);
                drl["liq_noe"] = reader1.GetValue(6).ToString().Trim();
                drl["estatus"] = reader1.GetValue(7).ToString().Trim();
                detliq.Rows.Add(drl);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT id_conse_liq, nom_desc FROM tb_cat_descuentos ORDER BY nom_desc";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                clavedesc.Add(reader1.GetValue(0).ToString().Trim());
                nombredesc.Add(reader1.GetValue(1).ToString().Trim());
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            for (int i = 0; i < nombredesc.Count; i++)
            {
                DDLDescuento.Items.Add(nombredesc[i]);
            }


            //cmnd1.CommandText = "SELECT hrp_recibo, hrp_num_unidades FROM tb_hist_recepcion WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
            //            " AND hrp_estatus <> 'C' AND lin_clave = '" + txt_lincve.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM'";
            //reader1 = cmnd1.ExecuteReader();
            //DataRow rw;
            //if (reader1.HasRows)
            //{
            //    while (reader1.Read())
            //    {
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "SELECT prov_clave FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + lbl_cveprov.Text + "'";


            //        cmnd1 = thisConnection.CreateCommand();
            //        rw = tbrecibos.NewRow();
            //        rw["hrp_recibo"] = reader1.GetValue(0).ToString().Trim();
            //        rw["hrp_cantidad"] = reader1.GetValue(1).ToString().Trim();
            //        tbrecibos.Rows.Add(rw);
            //    }
            //}
            //reader1.Close();
            //reader1.Dispose();
            //cmnd1.Dispose();
            thisConnection.Close();

            //foreach (DataRow r in tbrecibos.Rows)
            //{
            //    dtgRecibos.Rows.Add(r.ItemArray);
            //}

            historicoempaque();

            if (txtTL.Text == "PRO")
            {
                thisConnection.Open();
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT Ordp_Folio, Prod_cve, Prov_cve, Liq_Folio FROM Tb_Det_Liq_Planta WHERE Liq_Folio = '" + lbl_liquidacion.Text + "'";
                reader1 = cmnd1.ExecuteReader();
                DataRow rc;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        rc = recibs.NewRow();
                        rc["Ordp_Folio"] = reader1.GetValue(0).ToString().Trim();
                        rc["Prod_cve"] = reader1.GetValue(1).ToString().Trim();
                        rc["Prov_cve"] = reader1.GetValue(2).ToString().Trim();
                        rc["Liq_Folio"] = reader1.GetValue(3).ToString().Trim();
                        recibs.Rows.Add(rc);
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                thisConnection.Close();
            }

            if (txt_tipo.Text == "NACIONAL")
            {
                if (txtTipoLiq.Text == "consulta")
                    consultaliquidacionN();
                else
                    datosnuevonal();
            }
            if (txt_tipo.Text == "EXPORTACION")
            {
                if (txtTipoLiq.Text == "consulta")
                    consultaliquidacionE();
                else
                    datosnuevoexp();
                txt_tipocambio.Text = lblTipoCambio.Text;

            }

            if (txtTipoLiq.Text != "consulta")
            {
                descuentoautoservicio();
                calculoporcentaje();
                calculatotales();
            }




        }

        public void datosnuevonal()
        {
            thisConnection.Open();

            if (num_liq == "")
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT max(liq_folio) as liq_folio FROM tb_mstr_liquidacion";// ORDER BY liq_folio DESC";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    lbl_liquidacion.Text = Convert.ToString(Convert.ToInt32(reader1.GetValue(0).ToString().Trim()) + 1);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
            }


            string tp = "";
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT TOP 1 valor FROM tb_cat_tipocambio ORDER BY fecha DESC, dia DESC";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                tp = reader1.GetValue(0).ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();

            cmnd1.CommandText = "select sum(b.fcn_num_unidades) as cajas, SUM(b.fcn_precio_mn * b.fcn_num_unidades) AS importe, SUM(b.fcn_precio_usd * b.fcn_num_unidades) AS importe_usd" +
                " from tb_mstr_facturas_nal a, tb_det_facturas b" +
                " where (a.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND a.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                " and a.fcn_estatus <> 'C' and  b.fcn_folio = a.fcn_folio" +
                //" and b.fcn_tipo = a.fcn_lugar AND b.fcn_tipo <> 'EXP' AND b.prod_clave = '" + cveprod + "' /*and b.fcn_precio_mn > 0*/" +
                " and b.fcn_tipo = a.fcn_lugar AND a.um_clave = 'PESOS' AND b.prod_clave = '" + cveprod + "' /*and b.fcn_precio_mn > 0*/" +
                " AND a.fcn_monto <> a.ncr_monto" +
                " group by b.prod_clave" +
                " order by b.prod_clave";
            reader1 = cmnd1.ExecuteReader();
            DataRow tpprow;
            while (reader1.Read())
            {
                tpprow = tpp.NewRow();
                tpprow["fcn_folio"] = "";
                tpprow["fcn_tipo"] = "";
                tpprow["fcn_estatus"] = "";
                tpprow["unidades"] = reader1.GetDecimal(0);
                tpprow["precio_mn"] = (reader1.GetDecimal(1) > 0) ? reader1.GetDecimal(1) : (reader1.GetDecimal(2) * Convert.ToDecimal(tp));
                tpprow["lin_clave"] = "";
                tpprow["prod_clave"] = "";
                tpp.Rows.Add(tpprow);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            txt_tipocambio.Enabled = false;
            //txt_tipo.Text = tipo;
            chk_afecta.Checked = false;
            chk_afecta.Enabled = false;

            string prod_flejes = "";
            string prod_tarimas = "";
            string var_dec_prod_num_flejes = "0";
            string var_dec_prod_num_tarimas = "0";


            string var_chr_prod_flejes = "";
            string var_chr_prod_tarimas = "";
            string var_chr_prod_esquineros = "";
            string var_dec_prod_num_esquineros = "0";
            string var_chr_prod_nombre = "";
            string var_dec_enfriamiento = "0";
            string var_dec_prod_flete = "0";
            string var_dec_prod_comision = "0";

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prod_flejes, prod_tarimas, prod_num_flejes, prod_num_tarimas, prod_comision, prod_esquineros, prod_num_esquineros, prod_nombre, prod_enfriamiento, prod_flete FROM tb_cat_producto WHERE prod_clave = '" + this.cveprod + "' ORDER BY prod_clave";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    prod_flejes = reader1.GetValue(0).ToString().Trim();
                    prod_tarimas = reader1.GetValue(1).ToString().Trim();
                    var_dec_prod_num_flejes = reader1.GetValue(2).ToString().Trim();
                    var_dec_prod_num_tarimas = reader1.GetValue(3).ToString().Trim();
                    var_dec_prod_comision = reader1.GetValue(4).ToString().Trim();

                    var_chr_prod_flejes = reader1.GetValue(0).ToString().Trim();
                    var_chr_prod_tarimas = reader1.GetValue(1).ToString().Trim();
                    var_chr_prod_esquineros = reader1.GetValue(5).ToString().Trim();
                    var_dec_prod_num_esquineros = reader1.GetValue(6).ToString().Trim();
                    var_chr_prod_nombre = reader1.GetValue(7).ToString().Trim();
                    var_dec_enfriamiento = reader1.GetValue(8).ToString().Trim();
                    var_dec_prod_flete = reader1.GetValue(9).ToString().Trim();
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            if (Convert.ToDecimal(var_dec_prod_num_tarimas) == 0)
            {
                MessageBox.Show("No esta registrado el número de tarimas del producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                return;
            }
            if (Convert.ToDecimal(var_dec_prod_num_flejes) == 0)
            {
                MessageBox.Show("No esta registrado el número de cajas del producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                return;
            }

            lbl_cajas.Text = var_dec_prod_num_tarimas.ToString();
            lbl_flejes.Text = var_dec_prod_num_flejes.ToString();
            txt_valor_por.Text = var_dec_prod_comision.ToString();
            string lib = Convert.ToString((Convert.ToDecimal(neto_prod) / Convert.ToDecimal(cantidad)) * Convert.ToDecimal("2.2"));
            lbl_libras.Text = Convert.ToDecimal(lib).ToString("##0.0000");
            string var_dec_precio = "";
            string uno = "";
            string var_int_registros = "";
            string var_dec_unidades = "";

            var_dec_unidades = cantidad;

            foreach (DataRow rw in tpp.Rows)
            {
                uno = Math.Truncate(Convert.ToDecimal(rw["unidades"].ToString())).ToString();
                var_int_registros = Convert.ToInt32(uno).ToString();
                var_dec_precio = rw["precio_mn"].ToString();
            }

            if (Convert.ToDecimal((var_int_registros == "") ? "0" : var_int_registros) == 0)
            {
                var_dec_precio = "0";
            }
            else
            {
                var_dec_precio = Math.Round((Convert.ToDecimal(var_dec_precio) / Convert.ToDecimal(var_int_registros)), 2).ToString();
            }

            DataRow dtr = tcon.NewRow();
            dtr["cve_con"] = Convert.ToString(1);
            dtr["nombre_con"] = "Total de Cajas";
            dtr["unidades"] = Convert.ToDecimal(var_dec_unidades);
            dtr["precio"] = var_dec_precio;
            dtr["total"] = Convert.ToDecimal(var_dec_unidades) * Convert.ToDecimal(var_dec_precio);
            dtr["calculo"] = "1";
            tcon.Rows.Add(dtr);


            string var_chr_emp_clave = "";
            string var_chr_emp_nombre = "";
            decimal can = 0;
            decimal var_dec_precio_ultimo_empaque = 0;
            decimal tot = 0;
            decimal var_dec_empaque = 0;
            decimal var_dec_total = 0;

            cmnd1.CommandText = "SELECT T.emp_clave, T.comt_cantidad, E.emp_nombre FROM tb_mstr_comp_terminado T, tb_cat_empaques E WHERE T.emp_clave = E.emp_clave  AND T.prod_clave = '" + lbl_cveprod.Text + "' ORDER BY T.lin_clave, T.prod_clave, T.emp_clave"; //AND T.lin_clave = '" + txt_lincve.Text + "'
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                var_chr_emp_clave = reader1.GetValue(0).ToString().Trim();
                //cmnd2 = thisConnection.CreateCommand();
                //cmnd2.CommandText = "SELECT emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "'";
                //reader2 = cmnd2.ExecuteReader();
                //while (reader2.Read())
                //{
                var_chr_emp_nombre = reader1.GetValue(2).ToString().Trim();
                //}
                //reader2.Close();

                can = Convert.ToDecimal(var_dec_unidades) * reader1.GetDecimal(1);

                bool fnd = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_emp_clave + "'"))
                {
                    fnd = true;
                    var_dec_precio_ultimo_empaque = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_emp_clave + "'"))
                    {
                        fnd = true;
                        var_dec_precio_ultimo_empaque = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }

                if (var_dec_precio_ultimo_empaque == 0)
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "SELECT TOP 1 hrp_costo FROM tb_historico_recepcion WHERE emp_clave = '" + var_chr_emp_clave + "' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                    reader2 = cmnd2.ExecuteReader();
                    if (reader2.HasRows)
                    {
                        reader2.Read();
                        var_dec_precio_ultimo_empaque = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                    }
                    reader2.Close();
                    reader2.Dispose();
                    cmnd2.Dispose();

                    if (var_chr_emp_clave == "C0002")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0264' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "C0264")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0002' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "T0003")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0261' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "C0261")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'T0003' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }


                }
                else
                {
                    if (var_chr_emp_clave == "C0002")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0002")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0264")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0002"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0264"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);
                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;


                        ////DataSet ds = historicoempaque("C0264");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //////cmnd2 = thisConnection.CreateCommand();
                        //////cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0264' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        //////reader2 = cmnd2.ExecuteReader();
                        //////if (reader2.HasRows)
                        //////{
                        //////    precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        //////}
                        //////reader2.Close();
                        //////reader2.Dispose();
                        //////cmnd2.Dispose();

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0002", "C0264", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "C0264")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0264")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0002")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0264"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0002"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);

                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        ////DataSet ds = historicoempaque("C0002");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}
                        //////cmnd2 = thisConnection.CreateCommand();
                        //////cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0002' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        //////reader2 = cmnd2.ExecuteReader();
                        //////if (reader2.HasRows)
                        //////{
                        //////    precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        //////}
                        //////reader2.Close();
                        //////reader2.Dispose();
                        //////cmnd2.Dispose();

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0264", "C0002", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "T0003")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("T0003")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0261")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("T0003"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0261"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);
                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;

                        ////DataSet ds = historicoempaque("C0261");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}
                        //////cmnd2 = thisConnection.CreateCommand();
                        //////cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0261' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        //////reader2 = cmnd2.ExecuteReader();
                        //////if (reader2.HasRows)
                        //////{
                        //////    precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        //////}
                        //////reader2.Close();
                        //////reader2.Dispose();
                        //////cmnd2.Dispose();

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("T0003", "C0261", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "C0261")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0261")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("T0003")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0261"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("T0003"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if(precio_ult_caja_empaque_prom == 0) 
                        //    precio_final = precio_inicial_empaque_prom;
                        //else 
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);





                        ////DataSet ds = historicoempaque("T0003");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //////cmnd2 = thisConnection.CreateCommand();
                        //////cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) AS hrp_costo FROM tb_historico_recepcion WHERE emp_clave = 'T0003' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        //////reader2 = cmnd2.ExecuteReader();
                        //////if (reader2.HasRows)
                        //////{
                        //////    precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        //////}
                        //////reader2.Close();
                        //////reader2.Dispose();
                        //////cmnd2.Dispose();

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0261", "T0003", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;//Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                }
                //var_dec_precio_ultimo_empaque = Math.Round(fn_trae_precio_promedio_empaque(var_chr_emp_clave, this.f1, this.f2), 3);

                tot = Math.Round((Convert.ToDecimal(var_dec_unidades) * reader1.GetDecimal(1) * var_dec_precio_ultimo_empaque), 3);
                //var_dec_total = Math.Round((var_dec_total - tot), 3);

                DataRow dtr2 = tcon.NewRow();
                dtr2["cve_con"] = var_chr_emp_clave;
                dtr2["nombre_con"] = var_chr_emp_nombre;
                dtr2["unidades"] = can;
                dtr2["precio"] = Math.Round(var_dec_precio_ultimo_empaque, 3);
                dtr2["total"] = Math.Round(tot, 3) * -1;
                dtr2["calculo"] = "1";
                tcon.Rows.Add(dtr2);

                var_dec_empaque = var_dec_empaque + (can * var_dec_precio_ultimo_empaque);
            }
            reader1.Close();

            //TARIMAS
            DataRow dttarima = tcon.NewRow();
            dttarima["cve_con"] = "2";
            dttarima["nombre_con"] = "Tarimas";
            dttarima["unidades"] = Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3);
            decimal tot1 = 0;
            decimal tot2 = 0;
            decimal tot3 = 0;
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) AS total, hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
                "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C'" +
                "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND e.emp_nombre like 'TARIMA%'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    tot1 = tot1 + Convert.ToDecimal(reader1.GetValue(2).ToString().Trim());
                    tot2 = tot2 + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                }
            }
            reader1.Close();
            reader1.Dispose();
            decimal var_pu_emp = 0;
            if (tot1 == 0 || tot2 == 0)
            {
                //var_pu_emp = Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_tarimas, this.f1, this.f2), 3);
                //var_pu_emp = 0;
                //dtr3["precio"] = 0;
                //dtr3["total"] = 0;//Math.Round((var_pu_emp * Math.Round((var_dec_unidades / var_dec_prod_num_tarimas), 2)), 3) * -1;
                //tcon.Rows.Add(dtr3);
                bool fnd = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_tarimas + "'"))
                {
                    fnd = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_tarimas + "'"))
                    {
                        fnd = true;
                        var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }
            }
            else
            {
                tot3 = tot1 / tot2;
                var_pu_emp = Math.Round(tot3, 3);
            }

            dttarima["precio"] = var_pu_emp;
            dttarima["total"] = Math.Round((var_pu_emp * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 2)), 3) * -1;
            dttarima["calculo"] = "1";
            tcon.Rows.Add(dttarima);
            //FIN TARIMAS

            //ENFRIAMIENTO
            DataRow dtrenfria = tcon.NewRow();
            dtrenfria["cve_con"] = "3";
            dtrenfria["nombre_con"] = "Enfriamiento";
            dtrenfria["unidades"] = var_dec_unidades;
            var_pu_emp = Convert.ToDecimal(var_dec_enfriamiento);
            dtrenfria["precio"] = var_pu_emp;
            dtrenfria["total"] = Math.Round((Convert.ToDecimal(var_dec_unidades) * var_pu_emp), 3) * -1;
            tcon.Rows.Add(dtrenfria);
            //FIN ENFRIAMIENTO

            //FLEJES
            DataRow dtflejes = tcon.NewRow();
            dtflejes["cve_con"] = "4";
            //ceros redondeo
            decimal a1 = Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 0);
            dtflejes["nombre_con"] = "Flejes";
            dtflejes["unidades"] = Math.Round((Convert.ToDecimal(var_dec_prod_num_flejes) * a1), 3);//0
            //var_pu_emp = fn_trae_precio_promedio_empaque(var_chr_prod_flejes, this.f1, this.f2);
            bool fnd2 = false;
            foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_flejes + "'"))
            {
                fnd2 = true;
                var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
            }
            if (fnd2 == false)
            {
                foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_flejes + "'"))
                {
                    fnd2 = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                }
            }
            dtflejes["precio"] = var_pu_emp;
            dtflejes["total"] = Math.Round((Convert.ToDecimal(var_dec_prod_num_flejes) * a1), 3) * var_pu_emp * -1;//0
            tcon.Rows.Add(dtflejes);
            //FIN FLEJES

            //ESQUINEROS
            DataRow dtesq = tcon.NewRow();
            dtesq["cve_con"] = "5";
            dtesq["nombre_con"] = "Esquineros";
            dtesq["unidades"] = Math.Round(Convert.ToDecimal(var_dec_prod_num_esquineros) * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3), 3);


            tot1 = 0;
            tot2 = 0;
            tot3 = 0;
            cmnd1.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) AS total, hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
                "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C'" +
                "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND e.emp_nombre like 'ESQUINERO%'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    tot1 = tot1 + Convert.ToDecimal(reader1.GetValue(2).ToString().Trim());
                    tot2 = tot2 + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                }
            }
            reader1.Close();
            reader1.Dispose();
            decimal val1 = 0;
            decimal val2 = 0;
            if (tot1 == 0 || tot2 == 0)
            {
                //var_pu_emp = Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_esquineros, this.f1, this.f2), 3);
                bool fnd3 = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_esquineros + "'"))
                {
                    fnd3 = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd3 == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_esquineros + "'"))
                    {
                        fnd2 = true;
                        var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }
            }
            else
            {
                tot3 = tot1 / tot2;
                var_pu_emp = Math.Round(tot3, 3);//Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_esquineros, this.f1, this.f2), 3);

            }
            dtesq["precio"] = var_pu_emp;
            //decimal val1 = Math.Round(var_dec_prod_num_esquineros * Math.Round((var_dec_unidades / var_dec_prod_num_tarimas), 3), 3);
            //decimal val2 = var_pu_emp;
            val1 = Math.Round(Convert.ToDecimal(var_dec_prod_num_esquineros) * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3), 3);
            val2 = var_pu_emp;
            dtesq["total"] = Math.Round((val1 * val2), 3) * -1;
            dtesq["calculo"] = "1";
            tcon.Rows.Add(dtesq);
            //FIN ESQUINEROS

            ////FLETES
            if (lbl_cveprov.Text == "03" || lbl_cveprov.Text == "01" || lbl_cveprov.Text == "1328")
            {
                string mes = "";
                int anio = 0;
                var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //string mes2 = hoy.ToString("MMMM", CultureInfo.CreateSpecificCulture("en-US"));
                DateTimeFormatInfo fe1 = new CultureInfo("es-ES", false).DateTimeFormat;
                if (Convert.ToDateTime(f2) > Convert.ToDateTime(f1))
                    mes = fe1.GetMonthName(Convert.ToDateTime(f2).Month);
                else
                    mes = fe1.GetMonthName(Convert.ToDateTime(f1).Month);

                anio = Convert.ToDateTime(f1).Year;
                cmnd1.CommandText = "SELECT prod_clave, costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'NAL'  " +
                    "AND mes = '" + mes.ToUpper().ToString() + "' AND año = '" + anio + "' ORDER BY prod_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    if (reader1.GetValue(1).ToString().Trim() == "")
                    {
                        var_dec_precio_ultimo_empaque = 0;
                    }
                    else
                    {
                        var_dec_precio_ultimo_empaque = reader1.GetDecimal(1);
                    }

                }
                reader1.Close();

                DataRow dtfletes = tcon.NewRow();
                dtfletes["cve_con"] = "6";
                dtfletes["nombre_con"] = "Fletes";
                dtfletes["unidades"] = var_dec_unidades;
                dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
                dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                tcon.Rows.Add(dtfletes);
            }
            else
            {

                string mes = "";
                int anio = 0;
                var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //string mes2 = hoy.ToString("MMMM", CultureInfo.CreateSpecificCulture("en-US"));
                DateTimeFormatInfo fe1 = new CultureInfo("es-ES", false).DateTimeFormat;
                if (Convert.ToDateTime(f2) > Convert.ToDateTime(f1))
                    mes = fe1.GetMonthName(Convert.ToDateTime(f2).Month);
                else
                    mes = fe1.GetMonthName(Convert.ToDateTime(f1).Month);

                anio = Convert.ToDateTime(f1).Year;
                cmnd1.CommandText = "SELECT prod_clave, costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'NAL'  " +
                    "AND mes = '" + mes.ToUpper().ToString() + "' AND año = '" + anio + "' ORDER BY prod_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    if (reader1.GetValue(1).ToString().Trim() == "")
                    {
                        var_dec_precio_ultimo_empaque = 0;
                    }
                    else
                    {
                        var_dec_precio_ultimo_empaque = reader1.GetDecimal(1);
                    }

                }
                reader1.Close();

                DataRow dtfletes = tcon.NewRow();
                dtfletes["cve_con"] = "6";
                dtfletes["nombre_con"] = "Fletes";
                dtfletes["unidades"] = var_dec_unidades;
                dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
                dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                tcon.Rows.Add(dtfletes);

                //string sem_flt = "";
                //string ano_flt = "";

                ////consulta semana y año de la fecha inicial del rango 07/09/2021
                //string fch_dia_act = lbl_fecha1.Text;
                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT * FROM tb_cat_semanas WHERE '" + fch_dia_act + "' >= fecha1 AND '" + fch_dia_act + "' <= fecha2 ";
                //reader1 = cmnd1.ExecuteReader();
                //string sem_act = "";
                //string anio_act = "";
                //if (reader1.HasRows)
                //{
                //    reader1.Read();
                //    sem_act = reader1["semana"].ToString();
                //    anio_act = reader1["ano"].ToString();
                //}

                //sem_flt = sem_act;
                //ano_flt = anio_act;

                ////string f1 = "";
                ////string f2 = "";
                ////if (sem_act == "1")
                ////{
                ////    cmnd1 = thisConnection.CreateCommand();
                ////    cmnd1.CommandText = "select top 1 semana, ano, fecha1, fecha2 from tb_cat_semanas where ano = '" + (Convert.ToInt32(anio_act) - 1) + "' ORDER BY semana desc";
                ////    reader1 = cmnd1.ExecuteReader();
                ////    string sem_ant = "";
                ////    string anio_ant = "";
                ////    if (reader1.HasRows)
                ////    {
                ////        reader1.Read();
                ////        sem_ant = reader1["semana"].ToString();
                ////        anio_ant = reader1["ano"].ToString();
                ////    }
                ////}



                ////cmnd1 = thisConnection.CreateCommand();
                ////cmnd1.CommandText = "SELECT semana, ano FROM tb_cat_semanas WHERE fecha1 = '" + lbl_fecha1.Text + "' AND fecha2 = '" + lbl_fecha2.Text + "'";
                ////reader1 = cmnd1.ExecuteReader();
                ////if (reader1.HasRows)
                ////{
                ////    reader1.Read();
                ////    sem_flt = reader1["semana"].ToString().Trim();
                ////    ano_flt = reader1["ano"].ToString().Trim();
                ////}
                ////reader1.Close();
                ////cmnd1.ExecuteReader();
                ////cmnd1.Dispose();

                //var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //decimal cant_cajas_flt = 0;
                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT cajas, costo FROM Tb_Costos_FleteSem WHERE Semana = '" + sem_flt + "' AND Anio = '" + ano_flt + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND tipo = 'N'";
                //reader1 = cmnd1.ExecuteReader();
                //if (reader1.HasRows)
                //{
                //    while (reader1.Read())
                //    {
                //        var_dec_precio_ultimo_empaque = Convert.ToDecimal(reader1["costo"].ToString().Trim());// reader1.GetDecimal(1);
                //        cant_cajas_flt = Convert.ToDecimal(reader1["cajas"].ToString().Trim());

                //    }
                //}
                //else
                //{
                //    cant_cajas_flt = 0;
                //    var_dec_precio_ultimo_empaque = 0;
                //}

                //reader1.Close();

                //decimal c_cjs_liq = 0;//cajas reales a liquidar
                //decimal c_cajas = Convert.ToDecimal(var_dec_unidades);//cantidad a liquidar
                //if(c_cajas > cant_cajas_flt)
                //    c_cjs_liq = cant_cajas_flt;//cajas a liquidar > a cajas flete se toma cajas flete
                //else if (c_cajas < cant_cajas_flt)
                //    c_cjs_liq = c_cajas;//cajas a liquidar < cajas flete se toma cajas a liquidar
                //else
                //    c_cjs_liq = c_cajas;


                //DataRow dtfletes = tcon.NewRow();
                //dtfletes["cve_con"] = "6";
                //dtfletes["nombre_con"] = "Fletes";
                //dtfletes["unidades"] = c_cjs_liq;
                //dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                //dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(c_cjs_liq)), 3) * -1;
                //dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                //tcon.Rows.Add(dtfletes);
            }






            ////FIN FLETES

            //MERMAS Y RECLAMACIONES, NOTAS DE CREDITO, NOTAS DE CARGO, RECHAZOS POR CALIDAD

            tablanc.Clear();
            cmnd1 = thisConnection.CreateCommand();
            string mprov = "";
            mprov = lbl_cveprov.Text;
            cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + mprov + "' AND (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.clavep = '" + cveprod + "' and A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar and B.nc_estatus <> 'C' AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_nal = '0' " +
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.cveprov = '" + mprov + "' AND A.clavep = '" + cveprod + "' AND (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9803' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_nal = '0' --ORDER BY cveprov, prod_clave, fechap" +
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.cveprov = '" + mprov + "' AND A.clavep = '" + cveprod + "' AND (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9803' AND A.dnc_tipo in ('NCR', 'NCG') and A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar and B.nc_estatus <> 'C' AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_nal = '0'  " + //--ORDER BY cveprov, prod_clave, fechap
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.cveprov = '" + mprov + "' AND A.clavep = '" + cveprod + "' AND (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9813' AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_nal = '0'  " +//--ORDER BY cveprov, prod_clave, fechap
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.cveprov = '" + mprov + "' AND A.clavep = '" + cveprod + "' AND (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9814' AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar and B.nc_estatus <> 'C' AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_nal = '0'  ";//--ORDER BY cveprov, prod_clave, fechap
                                                                                                                                                                                                                                                                                                                                                                                                                                                         //" union" +
                                                                                                                                                                                                                                                                                                                                                                                                                                                         //" SELECT prod_nombre, nc_folio, dnc_cantidad, dnc_precio_mn, dnc_precio_usd, clavep, dnc_tipo, lin_clave, fechap FROM tb_det_notascyc" +
                                                                                                                                                                                                                                                                                                                                                                                                                                                         //" WHERE cveprov = '" + mprov + "' AND clavep = '" + cveprod + "' AND (fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND lin_clave = '9803' AND dnc_tipo = 'NCG' --ORDER BY cveprov, prod_clave, fechap";
            DataRow drnc;
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                drnc = tablanc.NewRow();
                drnc["prod_nombre"] = reader1.GetValue(0).ToString().Trim();
                drnc["nc_folio"] = reader1.GetValue(1).ToString().Trim();
                drnc["dnc_cantidad"] = reader1.GetValue(2).ToString().Trim();
                drnc["dnc_precio_mn"] = reader1.GetValue(3).ToString().Trim();
                drnc["dnc_precio_usd"] = reader1.GetValue(4).ToString().Trim();
                drnc["clavep"] = reader1.GetValue(5).ToString().Trim();
                drnc["dnc_tipo"] = reader1.GetValue(6).ToString().Trim();
                drnc["lin_clave"] = reader1.GetValue(7).ToString().Trim();
                tablanc.Rows.Add(drnc);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            decimal imp = 0;
            can = 0;
            tot = 0;
            DataRow dtmerma;
            if (tablanc.Rows.Count > 0)
            {
                bool ent = false;
                foreach (DataRow rnc in tablanc.Select("prod_nombre like '%MERMA%'"))
                {
                    ent = true;
                    can = can + Convert.ToDecimal(rnc["dnc_cantidad"].ToString());
                    //imp = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_mn"].ToString());
                    imp = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_mn"].ToString());
                    tot = tot + imp;
                    imp = 0;
                }

                if (ent == true)
                {
                    dtmerma = tcon.NewRow();
                    dtmerma["cve_con"] = "7";
                    dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                    dtmerma["unidades"] = can;
                    dtmerma["precio"] = Math.Round((tot / can), 3);
                    dtmerma["total"] = tot * -1;
                    dtmerma["calculo"] = "1";
                    tcon.Rows.Add(dtmerma);
                }
                else
                {
                    dtmerma = tcon.NewRow();
                    dtmerma["cve_con"] = "7";
                    dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                    dtmerma["unidades"] = Convert.ToDecimal("0.000");
                    dtmerma["precio"] = Convert.ToDecimal("0.000");
                    dtmerma["total"] = Convert.ToDecimal("0.000");
                    tcon.Rows.Add(dtmerma);
                }
            }
            else
            {
                dtmerma = tcon.NewRow();
                dtmerma["cve_con"] = "7";
                dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                dtmerma["unidades"] = Convert.ToDecimal("0.000");
                dtmerma["precio"] = Convert.ToDecimal("0.000");
                dtmerma["total"] = Convert.ToDecimal("0.000");
                tcon.Rows.Add(dtmerma);
            }



            can = 0;
            tot = 0;
            bool entracn = false;
            DataRow dtncr;
            if (tablanc.Rows.Count > 0)
            {
                foreach (DataRow rwnc in tablanc.Select("lin_clave = '9803' and dnc_tipo = 'NCR'"))
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_mn"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncr = tcon.NewRow();
                    dtncr["cve_con"] = "92";
                    dtncr["nombre_con"] = "Notas de Crédito x Dif. en Precio";
                    dtncr["unidades"] = can;
                    dtncr["precio"] = tot / can;
                    dtncr["total"] = tot * -1;
                    dtncr["calculo"] = "1";
                    tcon.Rows.Add(dtncr);
                }
                else
                {
                    //rechazos por calidad
                    dtncr = tcon.NewRow();
                    dtncr["cve_con"] = "8";
                    dtncr["nombre_con"] = "Rechazos por Calidad";
                    dtncr["unidades"] = Convert.ToDecimal("0.000");
                    dtncr["precio"] = Convert.ToDecimal("0.000");
                    dtncr["total"] = Convert.ToDecimal("0.000");
                    tcon.Rows.Add(dtncr);
                }
            }
            else
            {
                //rechazos por calidad
                dtncr = tcon.NewRow();
                dtncr["cve_con"] = "8";
                dtncr["nombre_con"] = "Rechazos por Calidad";
                dtncr["unidades"] = Convert.ToDecimal("0.000");
                dtncr["precio"] = Convert.ToDecimal("0.000");
                dtncr["total"] = Convert.ToDecimal("0.000");
                tcon.Rows.Add(dtncr);
            }
            entracn = false;

            can = 0;
            tot = 0;
            DataRow dtncg;
            if (tablanc.Rows.Count > 0)
            {
                DataTable dtvw = new DataTable();
                DataView dw = tablanc.DefaultView;
                dw.RowFilter = "lin_clave = '9803' and dnc_tipo = 'NCG'";
                dtvw = dw.ToTable();

                foreach (DataRow rwnc in dtvw.Rows)
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_mn"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncg = tcon.NewRow();
                    dtncg["cve_con"] = "93";
                    dtncg["nombre_con"] = "Notas de Cargo";
                    dtncg["unidades"] = can;
                    dtncg["precio"] = Math.Round(tot / can, 3);
                    dtncg["total"] = Math.Round(tot * -1, 3); //dtncg["total"] = Math.Round(tot * -1, 3);
                    dtncg["calculo"] = "1";
                    tcon.Rows.Add(dtncg);
                }

            }
            //FIN MERMAS Y RECLAMACIONES, NOTAS DE CREDITO, NOTAS DE CARGO, RECHAZOS POR CALIDAD

            //---Notas de credito y cargo por ACONDICIONAMIENTO DE EMPAQUE EN DESTINO 10/05/2021
            can = 0;
            tot = 0;
            entracn = false;
            DataRow dtncr2;
            if (tablanc.Rows.Count > 0)
            {
                foreach (DataRow rwnc in tablanc.Select("lin_clave = '9813' and dnc_tipo = 'NCR'"))
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_mn"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncr2 = tcon.NewRow();
                    dtncr2["cve_con"] = "108";
                    dtncr2["nombre_con"] = "Notas de Crédito x Acond. Emp. Destino";
                    dtncr2["unidades"] = can;
                    dtncr2["precio"] = Math.Round(tot / can, 3);
                    dtncr2["total"] = Math.Round(tot * -1, 3);
                    dtncr2["calculo"] = "1";
                    tcon.Rows.Add(dtncr2);

                    //var_dec_total = var_dec_total - Math.Round(tot, 3);
                }

            }
            entracn = false;

            can = 0;
            tot = 0;
            DataRow dtncg2;
            if (tablanc.Rows.Count > 0)
            {
                DataTable dtvw = new DataTable();
                DataView dw = tablanc.DefaultView;
                dw.RowFilter = "lin_clave = '9814' and dnc_tipo = 'NCG'";
                dtvw = dw.ToTable();

                foreach (DataRow rwnc in dtvw.Rows)
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_mn"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncg2 = tcon.NewRow();
                    dtncg2["cve_con"] = "107";
                    dtncg2["nombre_con"] = "Notas de Cargo x Acond. Emp. Origen";
                    dtncg2["unidades"] = can;
                    dtncg2["precio"] = Math.Round(tot / can, 3);
                    dtncg2["total"] = Math.Round(tot, 3);//dtncg["total"] = Math.Round(tot * -1, 3);
                    dtncg2["calculo"] = "1";
                    tcon.Rows.Add(dtncg2);
                }

            }
            //---Fin Notas de credito y cargo por ACONDICIONAMIENTO DE EMPAQUE EN DESTINO

            //AGREGAR CONCEPTO DE IMPORTE
            can = 0;
            tot = 0;
            DataRow dtrimporte = tcon.NewRow();
            dtrimporte["cve_con"] = "99";
            dtrimporte["nombre_con"] = "Impuesto";
            dtrimporte["unidades"] = Convert.ToDecimal(var_dec_unidades);
            dtrimporte["precio"] = 0;
            dtrimporte["total"] = Math.Round(tot * -1, 3);
            tcon.Rows.Add(dtrimporte);

            //PRESTAMOS
            //----------29/11/2017----------//
            foreach (DataRow rpre in dtPrestamos.Rows)
            {
                DataRow rwpre = tcon.NewRow();
                rwpre["cve_con"] = rpre["Lin_Clave"].ToString();
                rwpre["nombre_con"] = rpre["Descripcion_Art"].ToString() + " REF: " + rpre["Factura"].ToString();
                rwpre["unidades"] = "1";
                decimal res = 0;
                decimal res_moneda = 0;
                res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
                if (txt_tipo.Text == "NACIONAL")//NACIONAL
                {
                    if (rpre["Moneda"].ToString() == "DOLARES")
                        res_moneda = res * Convert.ToDecimal(lblTipoCambio.Text);
                    else
                        res_moneda = res;
                }
                if (txt_tipo.Text == "EXPORTACION")//EXPORTACION
                {
                    if (rpre["Moneda"].ToString() == "PESOS")
                        res_moneda = res / Convert.ToDecimal(txt_tipocambio.Text);
                    else
                        res_moneda = res;
                }
                rwpre["precio"] = Math.Round(res_moneda, 3);
                rwpre["total"] = Math.Round(res_moneda, 3) * -1;
                rwpre["valor"] = rpre["Id_Movimiento"].ToString();
                rwpre["moni"] = rpre["Moneda"].ToString();
                rwpre["saldo"] = Math.Round(res_moneda, 3);
                tcon.Rows.Add(rwpre);
            }
            //----------FIN 29/11/2017----------//
            ////can = 0;
            ////tot = 0;
            //foreach (DataRow rpre in dtPrestamos.Rows)
            //{
            //    DataRow rwpre = tcon.NewRow();
            //    rwpre["cve_con"] = rpre["Lin_Clave"].ToString();
            //    rwpre["nombre_con"] = rpre["Descripcion_Art"].ToString();
            //    rwpre["unidades"] = "1";
            //    decimal res = 0;
            //    res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
            //    rwpre["precio"] = Math.Round(res, 3);
            //    rwpre["total"] = Math.Round(res, 3) * -1;
            //    rwpre["valor"] = rpre["Id_Movimiento"].ToString();
            //    tcon.Rows.Add(rwpre);

            //    //bool fnd = lbl_producto.Text.Contains(rpre["Descripcion_Art"].ToString());
            //    //if (fnd == true)
            //    //{
            //    //DataRow rwpre = tcon.NewRow();
            //    //rwpre["cve_con"] = "95";
            //    //rwpre["nombre_con"] = "DESCUENTO PAGO ANTICIPADO";
            //    //rwpre["unidades"] = "1";
            //    //decimal res = 0;
            //    //res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
            //    //rwpre["precio"] = Math.Round(res, 3);
            //    //rwpre["total"] = Math.Round(res, 3) * -1;
            //    //tcon.Rows.Add(rwpre);
            //    //lblIdPrestamo.Text = rpre["Id_Movimiento"].ToString();
            //    //    break;
            //    //}
            //}
            //FIN PRESTAMOS

            //----------CALCULO DE MERMAS GAB----------
            can = 0;
            tot = 0;
            DataTable dtMermasGab = new DataTable();
            SqlDataAdapter adap = new SqlDataAdapter("spSISEMPLiquidacionesMermaGab", thisConnection);
            adap.SelectCommand.CommandType = CommandType.StoredProcedure;
            adap.SelectCommand.Parameters.AddWithValue("@FechaI", Convert.ToDateTime(f1).ToShortDateString());
            adap.SelectCommand.Parameters.AddWithValue("@FechaF", Convert.ToDateTime(f2).ToShortDateString());
            adap.SelectCommand.Parameters.AddWithValue("@Prod", cveprod);
            adap.Fill(dtMermasGab);
            decimal TCj = 0;
            decimal Flt = 0;
            decimal MyR = 0;
            decimal Com = 0;
            decimal TCo = 0;

            tcon.AsEnumerable()
                .Where(row => row.Field<string>("cve_con") == "1")
                .ToList()
                .ForEach(row =>
                {
                    TCj = Convert.ToDecimal(row["unidades"]);
                });

            tcon.AsEnumerable()
                .Where(row => row.Field<string>("cve_con") == "6")
                .ToList()
                .ForEach(row =>
                {
                    Flt = Math.Round((Convert.ToDecimal(row["total"]) / Convert.ToDecimal(row["unidades"]) * -1), 3);
                });

            tcon.AsEnumerable()
                .Where(row => row.Field<string>("cve_con") == "7")
                .ToList()
                .ForEach(row =>
                {
                    MyR = Math.Round((Convert.ToDecimal(row["total"]) * -1) / TCj, 3);
                });


            Com = Comision();

            TCo = Math.Round((Com * -1) / TCj, 3);

            Decimal tot_dec_merma = Flt + MyR + TCo;


            foreach (DataRow rt in dtMermasGab.Rows)
            {
                decimal cto_mem = Convert.ToDecimal(rt["Costo"].ToString()) - tot_dec_merma;
                tot = Convert.ToDecimal(rt["Cantidad"].ToString()) * cto_mem;
                can = Convert.ToDecimal(rt["Cantidad"].ToString());
                dtncg2 = tcon.NewRow();
                dtncg2["cve_con"] = "111";
                dtncg2["nombre_con"] = "Merma en Planta";
                dtncg2["unidades"] = can;
                dtncg2["precio"] = Math.Round(cto_mem, 3);//Math.Round(Convert.ToDecimal(rt["Costo"].ToString()), 3);//Math.Round(tot / can, 3);
                dtncg2["total"] = Math.Round(tot * -1, 3);//dtncg["total"] = Math.Round(tot * -1, 3);
                dtncg2["calculo"] = "1";
                tcon.Rows.Add(dtncg2);
            }
            //----------FIN CALCULO DE MERMAS GAB----------

            thisConnection.Close();

            for (int i = 0; i < tcon.Rows.Count; i++)
            {
                dtgConceptos.Rows.Add(tcon.Rows[i]["cve_con"].ToString(), tcon.Rows[i]["nombre_con"].ToString(), Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()).ToString("###,###,##0.000"),
                    Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("###,###,##0.000"), tcon.Rows[i]["valor"].ToString(), tcon.Rows[i]["moni"].ToString(), "", tcon.Rows[i]["saldo"].ToString(), tcon.Rows[i]["calculo"].ToString());
            }

            foreach (DataGridViewRow gr in dtgConceptos.Rows)
            {
                if (gr.Cells["clave"].Value.ToString().Length > 4)
                {
                    if (Convert.ToDecimal(gr.Cells["precio"].Value.ToString()) == 0)
                    {
                        gr.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
        }

        public void datosnuevoexp()
        {
            thisConnection.Open();

            if (num_liq == "")
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT max(liq_folio) as liq_folio FROM tb_mstr_liquidacion";// ORDER BY liq_folio DESC";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    lbl_liquidacion.Text = Convert.ToString(Convert.ToInt32(reader1.GetValue(0).ToString().Trim()) + 1);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
            }
            /*
             * [hrp_numliq] IN ('160328', '160325','160326','160327','160329','160330','160331','160332','160333')
             */


            cmnd1 = thisConnection.CreateCommand();

            //if (lbl_producto.Text.Contains("TOMATE") == true)
            //{
            //    cmnd1.CommandText = ";WITH TotalUnidadesPorFolio AS ( " +
            //                        "SELECT " +
            //                            "A.fcn_folio, " +
            //                            "SUM(A.fcn_num_unidades) AS total_unidades " +
            //                        "FROM " +
            //                            "tb_det_facturas A " +
            //                            "INNER JOIN tb_mstr_facturas_nal B ON A.fcn_folio = B.fcn_folio AND A.fcn_tipo = B.fcn_lugar " +
            //                        "where " +
            //                            "B.fcn_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
            //                            "AND B.fcn_estatus <> 'C' " +
            //                            "AND B.um_clave = 'USD' " +
            //                            "AND B.fcn_monto <> B.ncr_monto " +
            //                            "GROUP BY " +
            //                                "A.fcn_folio " +
            //                    ") " +
            //                    "SELECT " +
            //                        "b.prod_clave, " +
            //                        "SUM(b.fcn_num_unidades) AS cajas, " +
            //                        "SUM( " +
            //                        "(b.fcn_precio_usd - ROUND(ISNULL(a.fcn_monto_transporte / NULLIF(t.total_unidades, 0), 0), 2)) " +
            //                            "* b.fcn_num_unidades " +
            //                        ") AS importe_ajustado " +
            //                    "FROM " +
            //                        "tb_mstr_facturas_nal a " +
            //                    "JOIN " +
            //                        "tb_det_facturas b ON a.fcn_folio = b.fcn_folio AND b.fcn_tipo = a.fcn_lugar " +
            //                    "JOIN " +
            //                        "TotalUnidadesPorFolio t ON t.fcn_folio = a.fcn_folio " +
            //                    "WHERE " +
            //                        "a.fcn_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
            //                        "AND a.fcn_estatus <> 'C' " +
            //                        "AND a.um_clave = 'USD' " +
            //                        "AND b.prod_clave = '" + this.cveprod + "' " +
            //                        "AND a.fcn_monto <> a.ncr_monto " +
            //                    "GROUP BY " +
            //                        "b.prod_clave " +
            //                    "ORDER BY " +
            //                        "b.prod_clave;";

            //    reader1 = cmnd1.ExecuteReader();
            //    DataRow tppx;
            //    while (reader1.Read())
            //    {
            //        tppx = tppexp.NewRow();
            //        tppx["fcn_folio"] = "";
            //        tppx["fcn_tipo"] = "";
            //        tppx["fcn_estatus"] = "";
            //        tppx["unidades"] = (reader1.GetValue(1).ToString().Trim() == "") ? 0 : reader1.GetDecimal(1);
            //        tppx["precio_usd"] = (reader1.GetValue(2).ToString().Trim() == "") ? 0 : reader1.GetDecimal(2);
            //        tppx["lin_clave"] = "";
            //        tppx["prod_clave"] = "";
            //        tppexp.Rows.Add(tppx);
            //    }
            //    reader1.Close();
            //    reader1.Dispose();
            //    cmnd1.Dispose();
            //}
            //else 
            //{

            //}

            cmnd1.CommandText = "select sum(b.fcn_num_unidades) as cajas, SUM(b.fcn_precio_usd * b.fcn_num_unidades) AS importe" +
                    " from tb_mstr_facturas_nal a, tb_det_facturas b" +
                    " where (a.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND a.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') " +
                    //" and a.fcn_estatus <> 'C' and a.fcn_folio = b.fcn_folio and a.fcn_lugar = 'EXP'" +
                    " and a.fcn_estatus <> 'C' and a.fcn_folio = b.fcn_folio and a.um_clave = 'USD'" +
                    " and b.fcn_tipo = a.fcn_lugar AND b.prod_clave = '" + this.cveprod + "' AND a.fcn_monto <> a.ncr_monto --and b.fcn_precio_usd > 0 " +
                    " group by b.prod_clave" +
                    " order by b.prod_clave";

            reader1 = cmnd1.ExecuteReader();
            DataRow tppx;
            while (reader1.Read())
            {
                tppx = tppexp.NewRow();
                tppx["fcn_folio"] = "";
                tppx["fcn_tipo"] = "";
                tppx["fcn_estatus"] = "";
                tppx["unidades"] = (reader1.GetValue(0).ToString().Trim() == "") ? 0 : reader1.GetDecimal(0);
                tppx["precio_usd"] = (reader1.GetValue(1).ToString().Trim() == "") ? 0 : reader1.GetDecimal(1);
                tppx["lin_clave"] = "";
                tppx["prod_clave"] = "";
                tppexp.Rows.Add(tppx);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            //cmnd1.CommandText = "select sum(b.fcn_num_unidades) as cajas, SUM(b.fcn_precio_usd * b.fcn_num_unidades) AS importe" +
            //            " from tb_mstr_facturas_nal a, tb_det_facturas b" +
            //            " where (a.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND a.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') " +
            //            //" and a.fcn_estatus <> 'C' and a.fcn_folio = b.fcn_folio and a.fcn_lugar = 'EXP'" +
            //            " and a.fcn_estatus <> 'C' and a.fcn_folio = b.fcn_folio and a.um_clave = 'USD'" +
            //            " and b.fcn_tipo = a.fcn_lugar AND b.prod_clave = '" + this.cveprod + "' AND a.fcn_monto <> a.ncr_monto --and b.fcn_precio_usd > 0 " +
            //            " group by b.prod_clave" +
            //            " order by b.prod_clave";



            //txt_tipocambio.Enabled = true;
            //txt_tipo.Text = tipo;
            chk_afecta.Checked = false;
            chk_afecta.Enabled = false;

            string prod_flejes = "";
            string prod_tarimas = "";
            string var_dec_prod_num_flejes = "0";
            string var_dec_prod_num_tarimas = "0";
            string var_dec_prod_comision = "0";

            string var_chr_prod_flejes = "";
            string var_chr_prod_tarimas = "";
            string var_chr_prod_esquineros = "";
            string var_dec_prod_num_esquineros = "0";
            string var_chr_prod_nombre = "";
            string var_dec_enfriamiento = "0";
            string var_dec_prod_flete = "0";

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT prod_flejes, prod_tarimas, prod_num_flejes, prod_num_tarimas, prod_comision, prod_esquineros, prod_num_esquineros, prod_nombre, prod_enfriamiento, prod_flete FROM tb_cat_producto WHERE prod_clave = '" + this.cveprod + "' ORDER BY prod_clave";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    prod_flejes = reader1.GetValue(0).ToString().Trim();
                    prod_tarimas = reader1.GetValue(1).ToString().Trim();
                    var_dec_prod_num_flejes = reader1.GetValue(2).ToString().Trim();
                    var_dec_prod_num_tarimas = reader1.GetValue(3).ToString().Trim();
                    var_dec_prod_comision = reader1.GetValue(4).ToString().Trim();

                    var_chr_prod_flejes = reader1.GetValue(0).ToString().Trim();
                    var_chr_prod_tarimas = reader1.GetValue(1).ToString().Trim();
                    var_chr_prod_esquineros = reader1.GetValue(5).ToString().Trim();
                    var_dec_prod_num_esquineros = reader1.GetValue(6).ToString().Trim();
                    var_chr_prod_nombre = reader1.GetValue(7).ToString().Trim();
                    var_dec_enfriamiento = reader1.GetValue(8).ToString().Trim();
                    var_dec_prod_flete = reader1.GetValue(9).ToString().Trim();
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            if (Convert.ToDecimal(var_dec_prod_num_tarimas) == 0)
            {
                MessageBox.Show("No esta registrado el número de tarimas del producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                return;
            }
            if (Convert.ToDecimal(var_dec_prod_num_flejes) == 0)
            {
                MessageBox.Show("No esta registrado el número de flejes del producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                return;
            }

            lbl_cajas.Text = var_dec_prod_num_tarimas.ToString();
            lbl_flejes.Text = var_dec_prod_num_flejes.ToString();
            txt_valor_por.Text = var_dec_prod_comision.ToString();
            string lib = Convert.ToString((Convert.ToDecimal(neto_prod) / Convert.ToDecimal(cantidad)) * Convert.ToDecimal("2.2"));
            lbl_libras.Text = Convert.ToDecimal(lib).ToString("##0.000");
            string var_dec_precio = "";
            string uno = "";
            string var_int_registros = "0";
            string var_dec_unidades = "";

            var_dec_unidades = cantidad;

            foreach (DataRow rw in tppexp.Rows)
            {
                uno = Math.Truncate(Convert.ToDecimal(rw["unidades"].ToString())).ToString();
                var_int_registros = Convert.ToInt32(uno).ToString();
                var_dec_precio = rw["precio_usd"].ToString();
            }

            if (Convert.ToDecimal(var_int_registros) == 0)
            {
                var_dec_precio = "0";
            }
            else
            {
                var_dec_precio = Math.Round((Convert.ToDecimal(var_dec_precio) / Convert.ToDecimal(var_int_registros)), 2).ToString();
            }

            DataRow dtr = tcon.NewRow();
            dtr["cve_con"] = Convert.ToString(1);
            dtr["nombre_con"] = "Total de Cajas";
            dtr["unidades"] = Convert.ToDecimal(var_dec_unidades);
            dtr["precio"] = var_dec_precio;
            dtr["total"] = Convert.ToDecimal(var_dec_unidades) * Convert.ToDecimal(var_dec_precio);
            dtr["calculo"] = "1";
            tcon.Rows.Add(dtr);


            string var_chr_emp_clave = "";
            string var_chr_emp_nombre = "";
            decimal can = 0;
            decimal var_dec_precio_ultimo_empaque = 0;
            decimal tot = 0;
            decimal var_dec_empaque = 0;
            decimal var_dec_total = 0;

            cmnd1.CommandText = "SELECT T.emp_clave, T.comt_cantidad, E.emp_nombre FROM tb_mstr_comp_terminado T, tb_cat_empaques E WHERE T.emp_clave = E.emp_clave AND T.lin_clave = '" + txt_lincve.Text + "' AND T.prod_clave = '" + lbl_cveprod.Text + "' ORDER BY T.lin_clave, T.prod_clave, T.emp_clave";
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                var_chr_emp_clave = reader1.GetValue(0).ToString().Trim();
                //cmnd2 = thisConnection.CreateCommand();
                //cmnd2.CommandText = "SELECT emp_nombre FROM tb_cat_empaques WHERE emp_clave = '" + var_chr_emp_clave + "'";
                //reader2 = cmnd2.ExecuteReader();
                //while (reader2.Read())
                //{
                var_chr_emp_nombre = reader1.GetValue(2).ToString().Trim();
                //}
                //reader2.Close();

                can = Convert.ToDecimal(var_dec_unidades) * reader1.GetDecimal(1);

                bool fnd = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_emp_clave + "'"))
                {
                    fnd = true;
                    var_dec_precio_ultimo_empaque = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_emp_clave + "'"))
                    {
                        fnd = true;
                        var_dec_precio_ultimo_empaque = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }

                if (var_dec_precio_ultimo_empaque == 0)
                {
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "SELECT TOP 1 hrp_costo FROM tb_historico_recepcion WHERE emp_clave = '" + var_chr_emp_clave + "' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                    reader2 = cmnd2.ExecuteReader();
                    if (reader2.HasRows)
                    {
                        reader2.Read();
                        var_dec_precio_ultimo_empaque = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                    }
                    reader2.Close();
                    reader2.Dispose();
                    cmnd2.Dispose();

                    if (var_chr_emp_clave == "C0002")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0264' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "C0264")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0002' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "T0003")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'C0261' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                    else if (var_chr_emp_clave == "C0261")
                    {
                        decimal precio_ult_caja = 0;
                        decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT TOP 1 isnull(hrp_costo, 0) FROM tb_historico_recepcion WHERE emp_clave = 'T0003' AND hrp_estatus <> 'C' AND hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            precio_ult_caja = Convert.ToDecimal(reader2["hrp_costo"].ToString());
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        var_dec_precio_ultimo_empaque = Math.Round((precio_ult_caja + precio_ult_caja_comp) / 2, 3);
                    }
                }
                else
                {
                    if (var_chr_emp_clave == "C0002")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0002")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0264")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0002"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0264"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);

                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        ////DataSet ds = historicoempaque("C0264");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0002", "C0264", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "C0264")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0264")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0002")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0264"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0002"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);

                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;
                        ////DataSet ds = historicoempaque("C0002");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0264", "C0002", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "T0003")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("T0003")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0261")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("T0003"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("C0261"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);

                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;

                        ////DataSet ds = historicoempaque("C0261");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("T0003", "C0261", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                    else if (var_chr_emp_clave == "C0261")
                    {
                        //decimal precio_ult_caja_empaque1 = 0;
                        //decimal precio_ult_caja_empaque2 = 0;//var_dec_precio_ultimo_empaque;
                        //decimal precio_ult_caja_empaque_prom = 0;

                        //decimal precio_inicial_empaque1 = Convert.ToDecimal(recalculo_costos_empaque_inicial("C0261")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque2 = Convert.ToDecimal(recalculo_costos_empaque_inicial("T0003")); //CONSULTA DE COSTO INICIAL DEL MES
                        //decimal precio_inicial_empaque_prom = Math.Round((precio_inicial_empaque1 + precio_inicial_empaque2) / 2, 3); //PROMEDIO DE INICIAL DEL MES

                        //precio_ult_caja_empaque1 = Convert.ToDecimal(recalculo_costos_empaque("C0261"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO
                        //precio_ult_caja_empaque2 = Convert.ToDecimal(recalculo_costos_empaque("T0003"));//CONSULTA DE COSTO PROMEDIO SEGUN FECHA FINAL DE RANGO

                        ////SI ALGUNO DE LOS DOS EMPAQUES NO TIENE MOVIMIENTOS DE ENTRADAS NORMALES EL COSTO QUE SE TOMA ES EL DEL EMPAQUE QUE SI TUVO MOVIMIENTOS SIN REALIZAR NINGUN CALCULO
                        ////EN CASO CONTRARIO SI LOS DOS TUVIERON MOVIMIENTOS SE HACE EL CALCULO PARA CONOCER EL COSTO PROMEDIO
                        //if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque1;
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = precio_ult_caja_empaque2;
                        //else if (precio_ult_caja_empaque1 > 0 && precio_ult_caja_empaque2 > 0)
                        //    precio_ult_caja_empaque_prom = Math.Round((precio_ult_caja_empaque1 + precio_ult_caja_empaque2) / 2, 3);
                        //else if (precio_ult_caja_empaque1 == 0 && precio_ult_caja_empaque2 == 0)
                        //    precio_ult_caja_empaque_prom = 0;

                        //decimal precio_final = 0;
                        ////SI EL PRECIO SEGUN FECHA FINAL DE RANGO ES CERO ES PORQUE NINGUNO DE LOS DOS EMPAQUES TUVIERON ENTRADAS NORMALES EL PRECIO FINAL SERIA EL PROMEDIO DE LOS INICIALES
                        //if (precio_ult_caja_empaque_prom == 0)
                        //    precio_final = precio_inicial_empaque_prom;
                        //else
                        //    precio_final = Math.Round((precio_inicial_empaque_prom + precio_ult_caja_empaque_prom) / 2, 3);

                        ////decimal precio_ult_caja = 0;
                        ////decimal precio_ult_caja_comp = var_dec_precio_ultimo_empaque;

                        ////DataSet ds = historicoempaque("T0003");

                        ////bool fnd_2 = false;
                        ////foreach (DataRow rw in ds.Tables["historico"].Rows)
                        ////{
                        ////    fnd_2 = true;
                        ////    precio_ult_caja = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                        ////}
                        ////if (fnd_2 == false)
                        ////{
                        ////    foreach (DataRow rw in ds.Tables["catalogo"].Rows)
                        ////    {
                        ////        fnd_2 = true;
                        ////        precio_ult_caja = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                        ////    }
                        ////}

                        //29/08/2024
                        //Se reviso por parte de JAVIER y CLAUDIA el proceso por lo que el primer proceso planteado por Claudia era incorrecto. Javier explico que debiamos basarnos a las salidas y de alli sacar un costo
                        //promedio que saldría de la sumatorias de la cantidad y los montos totales
                        decimal precio_final = costo_promedio_caja_coliflor("C0261", "T0003", Convert.ToDateTime(f1).ToShortDateString(), Convert.ToDateTime(f2).ToShortDateString());
                        //fin 29/08/2024

                        var_dec_precio_ultimo_empaque = precio_final;
                    }
                }

                //var_dec_precio_ultimo_empaque = Math.Round(fn_trae_precio_promedio_empaque(var_chr_emp_clave, this.f1, this.f2), 3);

                tot = Math.Round((Convert.ToDecimal(var_dec_unidades) * reader1.GetDecimal(1) * var_dec_precio_ultimo_empaque), 3);
                //var_dec_total = Math.Round((var_dec_total - tot), 3);

                DataRow dtr2 = tcon.NewRow();
                dtr2["cve_con"] = var_chr_emp_clave;
                dtr2["nombre_con"] = var_chr_emp_nombre;
                dtr2["unidades"] = can;
                dtr2["precio"] = Math.Round(var_dec_precio_ultimo_empaque, 3);
                dtr2["total"] = Math.Round(tot, 3) * -1;
                dtr2["calculo"] = "1";
                tcon.Rows.Add(dtr2);

                var_dec_empaque = var_dec_empaque + (can * var_dec_precio_ultimo_empaque);
            }
            reader1.Close();

            //TARIMAS
            DataRow dttarima = tcon.NewRow();
            dttarima["cve_con"] = "2";
            dttarima["nombre_con"] = "Tarimas";
            dttarima["unidades"] = Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3);
            decimal tot1 = 0;
            decimal tot2 = 0;
            decimal tot3 = 0;
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) AS total, hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
                "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C'" +
                "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND e.emp_nombre like 'TARIMA%'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    tot1 = tot1 + Convert.ToDecimal(reader1.GetValue(2).ToString().Trim());
                    tot2 = tot2 + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                }
            }
            reader1.Close();
            reader1.Dispose();
            decimal var_pu_emp = 0;
            if (tot1 == 0 || tot2 == 0)
            {
                //var_pu_emp = Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_tarimas, this.f1, this.f2), 3);
                //var_pu_emp = 0;
                //dtr3["precio"] = 0;
                //dtr3["total"] = 0;//Math.Round((var_pu_emp * Math.Round((var_dec_unidades / var_dec_prod_num_tarimas), 2)), 3) * -1;
                //tcon.Rows.Add(dtr3);
                bool fnd = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_tarimas + "'"))
                {
                    fnd = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_tarimas + "'"))
                    {
                        fnd = true;
                        var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }
            }
            else
            {
                tot3 = tot1 / tot2;
                var_pu_emp = Math.Round(tot3, 3);
            }

            dttarima["precio"] = var_pu_emp;
            dttarima["total"] = Math.Round((var_pu_emp * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3)), 3) * -1;
            dttarima["calculo"] = "1";
            tcon.Rows.Add(dttarima);
            //FIN TARIMAS

            //ENFRIAMIENTO
            DataRow dtrenfria = tcon.NewRow();
            dtrenfria["cve_con"] = "3";
            dtrenfria["nombre_con"] = "Enfriamiento";
            dtrenfria["unidades"] = var_dec_unidades;
            var_pu_emp = Convert.ToDecimal(var_dec_enfriamiento);
            dtrenfria["precio"] = var_pu_emp;
            dtrenfria["total"] = Math.Round((Convert.ToDecimal(var_dec_unidades) * var_pu_emp), 3) * -1;
            tcon.Rows.Add(dtrenfria);
            //FIN ENFRIAMIENTO

            //FLEJES
            DataRow dtflejes = tcon.NewRow();
            dtflejes["cve_con"] = "4";
            dtflejes["nombre_con"] = "Flejes";
            //ceros redondeo
            decimal a1 = Math.Round(Convert.ToDecimal(var_dec_prod_num_flejes) * Math.Round(Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas), 3), 3);
            dtflejes["unidades"] = a1;//0
            //var_pu_emp = fn_trae_precio_promedio_empaque(var_chr_prod_flejes, this.f1, this.f2);
            bool fnd2 = false;
            foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_flejes + "'"))
            {
                fnd2 = true;
                var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
            }
            if (fnd2 == false)
            {
                foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_flejes + "'"))
                {
                    fnd2 = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                }
            }
            dtflejes["precio"] = var_pu_emp;
            dtflejes["total"] = Math.Round(a1 * var_pu_emp * -1, 3);//0
            tcon.Rows.Add(dtflejes);
            //FIN FLEJES

            //ESQUINEROS
            DataRow dtesq = tcon.NewRow();
            dtesq["cve_con"] = "5";
            dtesq["nombre_con"] = "Esquineros";
            dtesq["unidades"] = Math.Round(Convert.ToDecimal(var_dec_prod_num_esquineros) * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3), 3);


            tot1 = 0;
            tot2 = 0;
            tot3 = 0;
            cmnd1.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) AS total, hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
                "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C'" +
                "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND e.emp_nombre like 'ESQUINERO%'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    tot1 = tot1 + Convert.ToDecimal(reader1.GetValue(2).ToString().Trim());
                    tot2 = tot2 + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                }
            }
            reader1.Close();
            reader1.Dispose();
            decimal val1 = 0;
            decimal val2 = 0;
            if (tot1 == 0 || tot2 == 0)
            {
                //var_pu_emp = Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_esquineros, this.f1, this.f2), 3);
                bool fnd3 = false;
                foreach (DataRow rw in tppe.Select("emp_clave = '" + var_chr_prod_esquineros + "'"))
                {
                    fnd3 = true;
                    var_pu_emp = Math.Round(Convert.ToDecimal(rw["hrp_costo"].ToString()), 3);
                }
                if (fnd3 == false)
                {
                    foreach (DataRow rw in ecto.Select("emp_clave = '" + var_chr_prod_esquineros + "'"))
                    {
                        fnd2 = true;
                        var_pu_emp = Math.Round(Convert.ToDecimal(rw["emp_costo"].ToString()), 3);
                    }
                }
            }
            else
            {
                tot3 = tot1 / tot2;
                var_pu_emp = Math.Round(tot3, 3);//Math.Round(fn_trae_precio_promedio_empaque(var_chr_prod_esquineros, this.f1, this.f2), 3);

            }
            dtesq["precio"] = var_pu_emp;
            //decimal val1 = Math.Round(var_dec_prod_num_esquineros * Math.Round((var_dec_unidades / var_dec_prod_num_tarimas), 3), 3);
            //decimal val2 = var_pu_emp;
            val1 = Math.Round(Convert.ToDecimal(var_dec_prod_num_esquineros) * Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 3), 3);
            val2 = var_pu_emp;
            dtesq["total"] = Math.Round((val1 * val2), 3) * -1;
            dtesq["calculo"] = "1";
            tcon.Rows.Add(dtesq);
            //FIN ESQUINEROS

            //FLETES
            if (lbl_cveprov.Text == "01" || lbl_cveprov.Text == "03" || lbl_cveprov.Text == "1328")
            {
                string mes = "";
                int anio = 0;
                var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //string mes2 = hoy.ToString("MMMM", CultureInfo.CreateSpecificCulture("en-US"));
                DateTimeFormatInfo fe1 = new CultureInfo("es-ES", false).DateTimeFormat;
                if (Convert.ToDateTime(f2) > Convert.ToDateTime(f1))
                    mes = fe1.GetMonthName(Convert.ToDateTime(f2).Month);
                else
                    mes = fe1.GetMonthName(Convert.ToDateTime(f1).Month);

                anio = Convert.ToDateTime(f1).Year;
                cmnd1.CommandText = "SELECT prod_clave, costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'EXP'  " +
                    "AND mes = '" + mes.ToUpper().ToString() + "' AND año = '" + anio + "' ORDER BY prod_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    if (reader1.GetValue(1).ToString().Trim() == "")
                    {
                        var_dec_precio_ultimo_empaque = 0;
                    }
                    else
                    {
                        var_dec_precio_ultimo_empaque = reader1.GetDecimal(1);
                    }

                }
                reader1.Close();

                DataRow dtfletes = tcon.NewRow();
                dtfletes["cve_con"] = "6";
                dtfletes["nombre_con"] = "Fletes";
                dtfletes["unidades"] = var_dec_unidades;
                dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
                dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                tcon.Rows.Add(dtfletes);
            }
            else
            {
                string mes = "";
                int anio = 0;
                var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //string mes2 = hoy.ToString("MMMM", CultureInfo.CreateSpecificCulture("en-US"));
                DateTimeFormatInfo fe1 = new CultureInfo("es-ES", false).DateTimeFormat;
                if (Convert.ToDateTime(f2) > Convert.ToDateTime(f1))
                    mes = fe1.GetMonthName(Convert.ToDateTime(f2).Month);
                else
                    mes = fe1.GetMonthName(Convert.ToDateTime(f1).Month);

                anio = Convert.ToDateTime(f1).Year;
                cmnd1.CommandText = "SELECT prod_clave, costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'EXP'  " +
                    "AND mes = '" + mes.ToUpper().ToString() + "' AND año = '" + anio + "' ORDER BY prod_clave";
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    if (reader1.GetValue(1).ToString().Trim() == "")
                    {
                        var_dec_precio_ultimo_empaque = 0;
                    }
                    else
                    {
                        var_dec_precio_ultimo_empaque = reader1.GetDecimal(1);
                    }

                }
                reader1.Close();

                DataRow dtfletes = tcon.NewRow();
                dtfletes["cve_con"] = "6";
                dtfletes["nombre_con"] = "Fletes";
                dtfletes["unidades"] = var_dec_unidades;
                dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
                dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                tcon.Rows.Add(dtfletes);


                //string sem_flt = "";
                //string ano_flt = "";
                ////cmnd1 = thisConnection.CreateCommand();
                ////cmnd1.CommandText = "SELECT semana, ano FROM tb_cat_semanas WHERE fecha1 = '" + lbl_fecha1.Text + "' AND fecha2 = '" + lbl_fecha2.Text + "'";
                ////reader1 = cmnd1.ExecuteReader();
                ////if (reader1.HasRows)
                ////{
                ////    reader1.Read();
                ////    sem_flt = reader1["semana"].ToString().Trim();
                ////    ano_flt = reader1["ano"].ToString().Trim();
                ////}
                ////reader1.Close();
                ////cmnd1.ExecuteReader();
                ////cmnd1.Dispose();

                ////consulta semana y año de la fecha inicial del rango 07/09/2021
                //string fch_dia_act = lbl_fecha1.Text;
                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "SELECT * FROM tb_cat_semanas WHERE '" + fch_dia_act + "' >= fecha1 AND '" + fch_dia_act + "' <= fecha2 ";
                //reader1 = cmnd1.ExecuteReader();
                //string sem_act = "";
                //string anio_act = "";
                //if (reader1.HasRows)
                //{
                //    reader1.Read();
                //    sem_act = reader1["semana"].ToString();
                //    anio_act = reader1["ano"].ToString();
                //}

                //sem_flt = sem_act;
                //ano_flt = anio_act;

                //var_dec_precio_ultimo_empaque = Convert.ToDecimal("0.000");
                //decimal cant_cajas_flt = 0;
                //cmnd1 = thisConnection.CreateCommand();
                //// RCC 2 DE JUNIO 2023
                //// se detecto un error en esta consulta no se esta alimentado la variable ano_flt y no esta regresando los costos del flete correctamente
                //cmnd1.CommandText = "SELECT cajas, costo FROM Tb_Costos_FleteSem WHERE Semana = '" + sem_flt + "' AND Anio = '" + ano_flt + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND tipo = 'E'";
                //reader1 = cmnd1.ExecuteReader();
                //if (reader1.HasRows)
                //{
                //    while (reader1.Read())
                //    {
                //        var_dec_precio_ultimo_empaque = Convert.ToDecimal(reader1["costo"].ToString().Trim());// reader1.GetDecimal(1);
                //        cant_cajas_flt = Convert.ToDecimal(reader1["cajas"].ToString().Trim());
                //    }
                //}
                //else
                //{
                //    cant_cajas_flt = 0;
                //    var_dec_precio_ultimo_empaque = 0;
                //}

                //reader1.Close();

                //decimal c_cjs_liq = 0;//cajas reales a liquidar
                //decimal c_cajas = Convert.ToDecimal(var_dec_unidades);//cantidad a liquidar
                //if (c_cajas > cant_cajas_flt)
                //    c_cjs_liq = cant_cajas_flt;//cajas a liquidar > a cajas flete se toma cajas flete
                //else if (c_cajas < cant_cajas_flt)
                //    c_cjs_liq = c_cajas;//cajas a liquidar < cajas flete se toma cajas a liquidar
                //else
                //    c_cjs_liq = c_cajas;

                //DataRow dtfletes = tcon.NewRow();
                //dtfletes["cve_con"] = "6";
                //dtfletes["nombre_con"] = "Fletes";
                //dtfletes["unidades"] = c_cjs_liq;
                //dtfletes["precio"] = var_dec_precio_ultimo_empaque;
                //dtfletes["total"] = Math.Round((Convert.ToDecimal(var_dec_precio_ultimo_empaque) * Convert.ToDecimal(c_cjs_liq)), 3) * -1;
                //dtfletes["calculo"] = (var_dec_precio_ultimo_empaque == 0) ? "0" : "1";
                //tcon.Rows.Add(dtfletes);
            }

            //FIN FLETES

            //MERMAS Y RECLAMACIONES, NOTAS DE CREDITO, NOTAS DE CARGO, RECHAZOS POR CALIDAD

            tablanc.Clear();
            cmnd1 = thisConnection.CreateCommand();
            string mprov = "";
            mprov = lbl_cveprov.Text;
            cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + mprov + "' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.clavep = '" + lbl_cveprod.Text + "' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_exp = '0' " +
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE /*cveprov = '" + mprov + "' AND*/ A.clavep = '" + cveprod + "' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9803' AND A.dnc_tipo in ('NCR', 'NCG') and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_exp = '0' " +//--ORDER BY cveprov, prod_clave, fechap
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE /*cveprov = '" + mprov + "' AND*/ A.clavep = '" + cveprod + "' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9813' AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_exp = '0' " +//--ORDER BY cveprov, prod_clave, fechap
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE /*cveprov = '" + mprov + "' AND*/ A.clavep = '" + cveprod + "' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9814' AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_exp = '0' " +//--ORDER BY cveprov, prod_clave, fechap
                " union" +
                " SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B" +
                " WHERE /*cveprov = '" + mprov + "' AND*/ A.clavep = '" + cveprod + "' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9815' AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.liq_folio_exp = '0' ";//--ORDER BY cveprov, prod_clave, fechap

            //cmnd1.CommandText = "SELECT prod_nombre, nc_folio, dnc_cantidad, dnc_precio_mn, dnc_precio_usd, clavep, dnc_tipo, lin_clave, fechap FROM tb_det_notascyc" +
            //    " WHERE prod_nombre LIKE 'MERMA%' AND cveprov = '" + mprov + "' AND (fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND clavep = '" + cveprod + "'" +
            //    " union" +
            //    " SELECT prod_nombre, nc_folio, dnc_cantidad, dnc_precio_mn, dnc_precio_usd, clavep, dnc_tipo, lin_clave, fechap FROM tb_det_notascyc" +
            //    " WHERE cveprov = '" + mprov + "' AND clavep = '" + cveprod + "' AND (fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND lin_clave = '9803' --ORDER BY cveprov, prod_clave, fechap" +
            //    " union" +
            //    " SELECT prod_nombre, nc_folio, dnc_cantidad, dnc_precio_mn, dnc_precio_usd, clavep, dnc_tipo, lin_clave, fechap FROM tb_det_notascyc " +
            //    " WHERE cveprov = '" + mprov + "' AND clavep = '" + cveprod + "' AND (fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND lin_clave = '9803' AND dnc_tipo in ('NCR', 'NCG') --ORDER BY cveprov, prod_clave, fechap";
            //" union" +
            //" SELECT prod_nombre, nc_folio, dnc_cantidad, dnc_precio_mn, dnc_precio_usd, clavep, dnc_tipo, lin_clave, fechap FROM tb_det_notascyc" +
            //" WHERE cveprov = '" + mprov + "' AND clavep = '" + cveprod + "' AND (fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND lin_clave = '9803' AND dnc_tipo = 'NCG' --ORDER BY cveprov, prod_clave, fechap";
            DataRow drnc;
            reader1 = cmnd1.ExecuteReader();
            while (reader1.Read())
            {
                drnc = tablanc.NewRow();
                drnc["prod_nombre"] = reader1.GetValue(0).ToString().Trim();
                drnc["nc_folio"] = reader1.GetValue(1).ToString().Trim();
                drnc["dnc_cantidad"] = reader1.GetValue(2).ToString().Trim();
                drnc["dnc_precio_mn"] = reader1.GetValue(3).ToString().Trim();
                drnc["dnc_precio_usd"] = reader1.GetValue(4).ToString().Trim();
                drnc["clavep"] = reader1.GetValue(5).ToString().Trim();
                drnc["dnc_tipo"] = reader1.GetValue(6).ToString().Trim();
                drnc["lin_clave"] = reader1.GetValue(7).ToString().Trim();
                tablanc.Rows.Add(drnc);
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            decimal imp = 0;
            can = 0;
            tot = 0;
            DataRow dtmerma;
            if (tablanc.Rows.Count > 0)
            {
                bool ent = false;
                foreach (DataRow rnc in tablanc.Select("prod_nombre like '%MERMA%'"))
                {
                    ent = true;
                    can = can + Convert.ToDecimal(rnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                    //imp = can * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                }

                if (ent == true)
                {
                    dtmerma = tcon.NewRow();
                    dtmerma["cve_con"] = "7";
                    dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                    dtmerma["unidades"] = can;
                    dtmerma["precio"] = Math.Round((tot / can), 3);
                    dtmerma["total"] = tot * -1;
                    dtmerma["calculo"] = "1";
                    tcon.Rows.Add(dtmerma);
                }
                else
                {
                    dtmerma = tcon.NewRow();
                    dtmerma["cve_con"] = "7";
                    dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                    dtmerma["unidades"] = Convert.ToDecimal("0.000");
                    dtmerma["precio"] = Convert.ToDecimal("0.000");
                    dtmerma["total"] = Convert.ToDecimal("0.000");
                    tcon.Rows.Add(dtmerma);
                }
            }
            else
            {
                dtmerma = tcon.NewRow();
                dtmerma["cve_con"] = "7";
                dtmerma["nombre_con"] = "Mermas y Reclamaciones";
                dtmerma["unidades"] = Convert.ToDecimal("0.000");
                dtmerma["precio"] = Convert.ToDecimal("0.000");
                dtmerma["total"] = Convert.ToDecimal("0.000");
                tcon.Rows.Add(dtmerma);
            }

            //FITOSANITARIO BROCOLI Y COLIFLOR DE EXPORTACION
            DataRow dtfito;
            if (lbl_producto.Text.Contains("BROCCOLI") || lbl_producto.Text.Contains("BROCOLI") || lbl_producto.Text.Contains("COLIFLOR"))
            {
                dtfito = tcon.NewRow();
                dtfito["cve_con"] = "82";
                //ceros redondeo
                decimal ax = Math.Round((Convert.ToDecimal(var_dec_unidades) / Convert.ToDecimal(var_dec_prod_num_tarimas)), 0);
                dtfito["nombre_con"] = "FITOSANITARIO";
                dtfito["unidades"] = var_dec_unidades;//Math.Round((var_dec_prod_num_flejes * ax), 3);//0
                var_pu_emp = Convert.ToDecimal("0.06");//fn_trae_precio_promedio_empaque(var_chr_prod_flejes, this.f1, this.f2);
                dtfito["precio"] = var_pu_emp;
                //dtfito["total"] = Math.Round((Convert.ToDecimal(var_dec_unidades) * ax), 3) * var_pu_emp * -1;//0
                dtfito["total"] = Math.Round(Convert.ToDecimal(var_dec_unidades), 3) * var_pu_emp * -1;//0
                tcon.Rows.Add(dtfito);

                //var_dec_total = var_dec_total - (var_pu_emp * Convert.ToDecimal(var_dec_unidades));
            }
            //FIN FITOSANITARIO BROCOLI Y COLIFLOR DE EXPORTACION

            //FUMIGACIONES
            decimal tot1x = 0;
            decimal tot2x = 0;
            decimal tot3x = 0;
            if (lbl_producto.Text.Contains("APIO") || lbl_producto.Text.Contains("KALE") || (lbl_producto.Text.Contains("LECHUGA") && lbl_producto.Text.Contains("OREJONA")) || lbl_producto.Text.Contains("SWISS CHARD") || lbl_producto.Text.Contains("RAINBOW CHARD"))
            {
                bool band = true;
                if (lbl_cveprod.Text == "09009ESK28")
                    band = false;
                if (lbl_cveprod.Text == "16KAML1220")
                    band = false;
                if (lbl_cveprod.Text == "16KAOML121")
                    band = false;
                if (lbl_cveprod.Text == "16KAORCH42")
                    band = false;

                if (lbl_cveprod.Text == "09HOLEOR33")
                    band = false;
                if (lbl_cveprod.Text == "09HOORML25")
                    band = false;
                if (lbl_cveprod.Text == "09HOLEOR25")
                    band = false;
                if (lbl_cveprod.Text == "16001HLO12")
                    band = false;
                if (lbl_cveprod.Text == "09LEROCH62")
                    band = false;
                if (lbl_cveprod.Text == "09009LOH14")
                    band = false;
                if (lbl_cveprod.Text == "09009LEO41")
                    band = false;
                if (lbl_cveprod.Text == "09009LEO45")
                    band = false;
                if (lbl_cveprod.Text == "09009LOB62")
                    band = false;
                if (lbl_cveprod.Text == "05005LO1X4")
                    band = false;
                if (lbl_cveprod.Text == "05005LETOR")
                    band = false;
                if (lbl_cveprod.Text == "05005LETAY")
                    band = false;
                if (lbl_cveprod.Text == "09TALEOJ41")
                    band = false;
                if (band == true)
                {
                    bool fnd = false;
                    cmnd2 = thisConnection.CreateCommand();
                    cmnd2.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) as total, h.hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
                        "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
                        "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND h.emp_clave in ('M2628', 'N3742')";
                    reader2 = cmnd2.ExecuteReader();
                    if (reader2.HasRows)
                    {
                        while (reader2.Read())
                        {
                            tot1x = tot1x + Convert.ToDecimal(reader2.GetValue(2).ToString().Trim());
                            //tot2x = tot2x + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
                        }
                        fnd = true;
                        //tot3x = tot1x / tot2x;

                        //TOTAL DE VENTAS
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%APIO%' AND DF.lin_clave in ('01', '16') AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY DF.lin_clave, DF.prod_clave " +
                            "ORDER BY DF.lin_clave, DF.prod_clave";
                        reader3 = cmnd3.ExecuteReader();
                        while (reader3.Read())
                        {
                            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }

                        //TOTAL DE VENTAS
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%KALE%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY DF.lin_clave, DF.prod_clave " +
                            "ORDER BY DF.lin_clave, DF.prod_clave";
                        reader3 = cmnd3.ExecuteReader();
                        while (reader3.Read())
                        {
                            if (reader3.GetValue(1).ToString().Trim() == "09009ESK28")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "16KAML1220")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "16KAOML121")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "16KAORCH42")
                                continue;
                            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }

                        //TOTAL DE VENTAS
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%LECHUGA%OREJONA%' AND DF.lin_clave IN ('05', '09', '16') AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY DF.lin_clave, DF.prod_clave " +
                            "ORDER BY DF.lin_clave, DF.prod_clave";
                        reader3 = cmnd3.ExecuteReader();
                        while (reader3.Read())
                        {
                            if (reader3.GetValue(1).ToString().Trim() == "09HOLEOR33")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09HOORML25")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09HOLEOR25")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "16001HLO12")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09LEROCH62")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09009LOH14")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09009LEO41")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09009LEO45")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09009LOB62")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "05005LO1X4")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "05005LETOR")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "05005LETAY")
                                continue;
                            if (reader3.GetValue(1).ToString().Trim() == "09TALEOJ41")
                                continue;
                            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }

                        //TOTAL DE VENTAS
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%SWISS%CHARD%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                            "GROUP BY DF.lin_clave, DF.prod_clave " +
                            "ORDER BY DF.lin_clave, DF.prod_clave";
                        reader3 = cmnd3.ExecuteReader();
                        while (reader3.Read())
                        {
                            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }

                        //TOTAL DE VENTAS
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%RAINBOW%CHARD%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                            "GROUP BY DF.lin_clave, DF.prod_clave " +
                            "ORDER BY DF.lin_clave, DF.prod_clave";
                        reader3 = cmnd3.ExecuteReader();
                        while (reader3.Read())
                        {
                            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }
                    }
                    reader2.Close();
                    reader2.Dispose();
                    cmnd2.Dispose();

                    if (tot1x == 0 || tot2x == 0)
                    {
                        tot3x = precioempaque(Convert.ToDateTime(f2));
                    }
                    else
                    {
                        tot3x = tot1x / tot2x;
                        //dtr3 = tcon.NewRow();
                        //dtr3["cve_con"] = "83";
                        //dtr3["nombre_con"] = "FUMIGACIONES";
                        //dtr3["unidades"] = var_dec_unidades;
                        //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
                        //dtr3["total"] = (tot3x * var_dec_unidades) * -1;//tot * -1;
                        //tcon.Rows.Add(dtr3);
                    }

                    //if (fnd == true)
                    //{
                    DataRow dttfum;
                    dttfum = tcon.NewRow();
                    dttfum["cve_con"] = "83";
                    dttfum["nombre_con"] = "FUMIGACIONES";
                    dttfum["unidades"] = var_dec_unidades;
                    dttfum["precio"] = Math.Round(tot3x, 3);//tot / can;
                    dttfum["total"] = (tot3x * Convert.ToDecimal(var_dec_unidades)) * -1;
                    dttfum["calculo"] = "1";
                    tcon.Rows.Add(dttfum);
                    //}
                    //var_dec_total = var_dec_total - (tot3x * var_dec_unidades);
                }

            }

            //if (lbl_producto.Text.Contains("KALE"))
            //{
            //    bool band = true;
            //    if (lbl_cveprod.Text == "09009ESK28")
            //        band = false;
            //    if (lbl_cveprod.Text == "16KAML1220")
            //        band = false;
            //    if (lbl_cveprod.Text == "16KAOML121")
            //        band = false;
            //    if (lbl_cveprod.Text == "16KAORCH42")
            //        band = false;
            //    if (band == true)
            //    {
            //        bool fnd = false;
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) as total, h.hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
            //            "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
            //            "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND h.emp_clave in ('M2628', 'N3742')";
            //        reader2 = cmnd2.ExecuteReader();
            //        if (reader2.HasRows)
            //        {
            //            while (reader2.Read())
            //            {
            //                tot1x = tot1x + Convert.ToDecimal(reader2.GetValue(2).ToString().Trim());
            //                //tot2x = tot2x + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
            //            }
            //            fnd = true;
            //            //tot3x = tot1x / tot2x;

            //            //TOTAL DE VENTAS
            //            cmnd3 = thisConnection.CreateCommand();
            //            cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
            //                "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
            //                "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%KALE%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
            //                "GROUP BY DF.lin_clave, DF.prod_clave " +
            //                "ORDER BY DF.lin_clave, DF.prod_clave";
            //            reader3 = cmnd3.ExecuteReader();
            //            while (reader3.Read())
            //            {
            //                if (reader3.GetValue(1).ToString().Trim() == "09009ESK28")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "16KAML1220")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "16KAOML121")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "16KAORCH42")
            //                    continue;
            //                tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
            //            }

            //        }
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();

            //        if (tot1x == 0 || tot2x == 0)
            //        {
            //            tot3x = precioempaque(Convert.ToDateTime(f2));
            //        }
            //        else
            //        {
            //            tot3x = tot1x / tot2x;
            //            //dtr3 = tcon.NewRow();
            //            //dtr3["cve_con"] = "83";
            //            //dtr3["nombre_con"] = "FUMIGACIONES";
            //            //dtr3["unidades"] = var_dec_unidades;
            //            //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //            //dtr3["total"] = (tot3x * var_dec_unidades) * -1;//tot * -1;
            //            //tcon.Rows.Add(dtr3);
            //        }

            //        //if (fnd == true)
            //        //{
            //        //dtr3 = tcon.NewRow();
            //        //dtr3["cve_con"] = "83";
            //        //dtr3["nombre_con"] = "FUMIGACIONES";
            //        //dtr3["unidades"] = var_dec_unidades;
            //        //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //        //dtr3["total"] = (tot3x * var_dec_unidades) * -1;
            //        //tcon.Rows.Add(dtr3);
            //        DataRow dttfum;
            //        dttfum = tcon.NewRow();
            //        dttfum["cve_con"] = "83";
            //        dttfum["nombre_con"] = "FUMIGACIONES";
            //        dttfum["unidades"] = var_dec_unidades;
            //        dttfum["precio"] = Math.Round(tot3x, 3);//tot / can;
            //        dttfum["total"] = (tot3x * Convert.ToDecimal(var_dec_unidades)) * -1;
            //        tcon.Rows.Add(dttfum);
            //        //}
            //        //var_dec_total = var_dec_total - (tot3x * var_dec_unidades);
            //    }

            //}

            //if (lbl_producto.Text.Contains("LECHUGA") && lbl_producto.Text.Contains("OREJONA"))
            //{
            //    bool band = true;
            //    if (lbl_cveprod.Text == "09HOLEOR33")
            //        band = false;
            //    if (lbl_cveprod.Text == "09HOORML25")
            //        band = false;
            //    if (lbl_cveprod.Text == "09HOLEOR25")
            //        band = false;
            //    if (lbl_cveprod.Text == "16001HLO12")
            //        band = false;
            //    if (lbl_cveprod.Text == "09LEROCH62")
            //        band = false;
            //    if (lbl_cveprod.Text == "09009LOH14")
            //        band = false;
            //    if (lbl_cveprod.Text == "09009LEO41")
            //        band = false;
            //    if (lbl_cveprod.Text == "09009LEO45")
            //        band = false;
            //    if (lbl_cveprod.Text == "09009LOB62")
            //        band = false;
            //    if (lbl_cveprod.Text == "05005LO1X4")
            //        band = false;
            //    if (lbl_cveprod.Text == "05005LETOR")
            //        band = false;
            //    if (lbl_cveprod.Text == "05005LETAY")
            //        band = false;
            //    if (lbl_cveprod.Text == "09TALEOJ41")
            //        band = false;
            //    if (band == true)
            //    {
            //        bool fnd = false;
            //        cmnd2 = thisConnection.CreateCommand();
            //        cmnd2.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) as total, h.hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
            //            "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
            //            "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND h.emp_clave in ('M2628', 'N3742')";
            //        reader2 = cmnd2.ExecuteReader();
            //        if (reader2.HasRows)
            //        {
            //            while (reader2.Read())
            //            {
            //                tot1x = tot1x + Convert.ToDecimal(reader2.GetValue(2).ToString().Trim());
            //                //tot2x = tot2x + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
            //            }
            //            fnd = true;
            //            //tot3x = tot1x / tot2x;

            //            //TOTAL DE VENTAS
            //            cmnd3 = thisConnection.CreateCommand();
            //            cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
            //                "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
            //                "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%LECHUGA%OREJONA%' AND DF.lin_clave IN ('05', '09', '16') AND DF.fcn_tipo = F.fcn_lugar " +
            //                "GROUP BY DF.lin_clave, DF.prod_clave " +
            //                "ORDER BY DF.lin_clave, DF.prod_clave";
            //            reader3 = cmnd3.ExecuteReader();
            //            while (reader3.Read())
            //            {
            //                if (reader3.GetValue(1).ToString().Trim() == "09HOLEOR33")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09HOORML25")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09HOLEOR25")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "16001HLO12")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09LEROCH62")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09009LOH14")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09009LEO41")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09009LEO45")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09009LOB62")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "05005LO1X4")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "05005LETOR")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "05005LETAY")
            //                    continue;
            //                if (reader3.GetValue(1).ToString().Trim() == "09TALEOJ41")
            //                    continue;
            //                tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
            //            }

            //        }
            //        reader2.Close();
            //        reader2.Dispose();
            //        cmnd2.Dispose();

            //        if (tot1x == 0 || tot2x == 0)
            //        {
            //            tot3x = precioempaque(Convert.ToDateTime(f2));
            //        }
            //        else
            //        {
            //            tot3x = tot1x / tot2x;
            //            //dtr3 = tcon.NewRow();
            //            //dtr3["cve_con"] = "83";
            //            //dtr3["nombre_con"] = "FUMIGACIONES";
            //            //dtr3["unidades"] = var_dec_unidades;
            //            //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //            //dtr3["total"] = (tot3x * var_dec_unidades) * -1;//tot * -1;
            //            //tcon.Rows.Add(dtr3);
            //        }

            //        //if (fnd == true)
            //        //{
            //        //dtr3 = tcon.NewRow();
            //        //dtr3["cve_con"] = "83";
            //        //dtr3["nombre_con"] = "FUMIGACIONES";
            //        //dtr3["unidades"] = var_dec_unidades;
            //        //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //        //dtr3["total"] = (tot3x * var_dec_unidades) * -1;
            //        //tcon.Rows.Add(dtr3);

            //        DataRow dttfum;
            //        dttfum = tcon.NewRow();
            //        dttfum["cve_con"] = "83";
            //        dttfum["nombre_con"] = "FUMIGACIONES";
            //        dttfum["unidades"] = var_dec_unidades;
            //        dttfum["precio"] = Math.Round(tot3x, 3);//tot / can;
            //        dttfum["total"] = (tot3x * Convert.ToDecimal(var_dec_unidades)) * -1;
            //        tcon.Rows.Add(dttfum);
            //        //}
            //        //var_dec_total = var_dec_total - (tot3x * var_dec_unidades);
            //    }

            //}

            //if (lbl_producto.Text.Contains("SWISS CHARD"))
            //{
            //    bool fnd = false;
            //    cmnd2 = thisConnection.CreateCommand();
            //    cmnd2.CommandText = "select h.emp_clave, e.emp_nombre, (h.hrp_cantidad * h.hrp_costo) as total, h.hrp_cantidad from tb_historico_recepcion h, tb_cat_empaques e " +
            //        "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
            //        "and h.hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and h.hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND h.emp_clave in ('M2628', 'N3742')";
            //    reader2 = cmnd2.ExecuteReader();
            //    if (reader2.HasRows)
            //    {
            //        while (reader2.Read())
            //        {
            //            tot1x = tot1x + Convert.ToDecimal(reader2.GetValue(2).ToString().Trim());
            //            //tot2x = tot2x + Convert.ToDecimal(reader1.GetValue(3).ToString().Trim());
            //        }
            //        fnd = true;
            //        //tot3x = tot1x / tot2x;

            //        //TOTAL DE VENTAS
            //        cmnd3 = thisConnection.CreateCommand();
            //        cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
            //            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
            //            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%SWISS%CHARD%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
            //            "GROUP BY DF.lin_clave, DF.prod_clave " +
            //            "ORDER BY DF.lin_clave, DF.prod_clave";
            //        reader3 = cmnd3.ExecuteReader();
            //        while (reader3.Read())
            //        {
            //            tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
            //        }

            //    }
            //    reader2.Close();
            //    reader2.Dispose();
            //    cmnd2.Dispose();

            //    if (tot1x == 0 || tot2x == 0)
            //    {
            //        tot3x = precioempaque(Convert.ToDateTime(f2));
            //    }
            //    else
            //    {
            //        tot3x = tot1x / tot2x;
            //        //dtr3 = tcon.NewRow();
            //        //dtr3["cve_con"] = "83";
            //        //dtr3["nombre_con"] = "FUMIGACIONES";
            //        //dtr3["unidades"] = var_dec_unidades;
            //        //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //        //dtr3["total"] = (tot3x * var_dec_unidades) * -1;//tot * -1;
            //        //tcon.Rows.Add(dtr3);
            //    }

            //    //if (fnd == true)
            //    //{
            //    //dtr3 = tcon.NewRow();
            //    //dtr3["cve_con"] = "83";
            //    //dtr3["nombre_con"] = "FUMIGACIONES";
            //    //dtr3["unidades"] = var_dec_unidades;
            //    //dtr3["precio"] = Math.Round(tot3x, 3);//tot / can;
            //    //dtr3["total"] = (tot3x * var_dec_unidades) * -1;
            //    //tcon.Rows.Add(dtr3);

            //    DataRow dttfum;
            //    dttfum = tcon.NewRow();
            //    dttfum["cve_con"] = "83";
            //    dttfum["nombre_con"] = "FUMIGACIONES";
            //    dttfum["unidades"] = var_dec_unidades;
            //    dttfum["precio"] = Math.Round(tot3x, 3);//tot / can;
            //    dttfum["total"] = (tot3x * Convert.ToDecimal(var_dec_unidades)) * -1;
            //    tcon.Rows.Add(dttfum);
            //    //}
            //    //var_dec_total = var_dec_total - (tot3x * var_dec_unidades);
            //}

            //FIN FUMIGACIONES

            can = 0;
            tot = 0;
            bool entracn = false;
            DataRow dtncr;
            if (tablanc.Rows.Count > 0)
            {
                foreach (DataRow rwnc in tablanc.Select("lin_clave = '9803' and dnc_tipo = 'NCR'"))
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncr = tcon.NewRow();
                    dtncr["cve_con"] = "92";
                    dtncr["nombre_con"] = "Notas de Crédito x Dif. en Precio";
                    dtncr["unidades"] = can;
                    dtncr["precio"] = Math.Round(tot / can, 3);
                    dtncr["total"] = Math.Round(tot * -1, 3);
                    dtncr["calculo"] = "1";
                    tcon.Rows.Add(dtncr);

                    //var_dec_total = var_dec_total - Math.Round(tot, 3);
                }

            }
            entracn = false;

            can = 0;
            tot = 0;
            DataRow dtncg;
            if (tablanc.Rows.Count > 0)
            {
                DataTable dtvw = new DataTable();
                DataView dw = tablanc.DefaultView;
                dw.RowFilter = "lin_clave = '9803' and dnc_tipo = 'NCG'";
                dtvw = dw.ToTable();

                foreach (DataRow rwnc in dtvw.Rows)
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncg = tcon.NewRow();
                    dtncg["cve_con"] = "93";
                    dtncg["nombre_con"] = "Notas de Cargo";
                    dtncg["unidades"] = can;
                    dtncg["precio"] = Math.Round(tot / can, 3);
                    dtncg["total"] = Math.Round(tot, 3);//dtncg["total"] = Math.Round(tot * -1, 3);
                    dtncg["calculo"] = "1";
                    tcon.Rows.Add(dtncg);
                }

            }


            //---Notas de credito y cargo por ACONDICIONAMIENTO DE EMPAQUE EN DESTINO 10/05/2021
            can = 0;
            tot = 0;
            entracn = false;
            DataRow dtncr2;
            if (tablanc.Rows.Count > 0)
            {
                foreach (DataRow rwnc in tablanc.Select("lin_clave = '9813' and dnc_tipo = 'NCR'"))
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncr2 = tcon.NewRow();
                    dtncr2["cve_con"] = "108";
                    dtncr2["nombre_con"] = "Notas de Crédito x Acond. Emp. Destino";
                    dtncr2["unidades"] = can;
                    dtncr2["precio"] = Math.Round(tot / can, 3);
                    dtncr2["total"] = Math.Round(tot * -1, 3);
                    dtncr2["calculo"] = "1";
                    tcon.Rows.Add(dtncr2);

                    //var_dec_total = var_dec_total - Math.Round(tot, 3);
                }

            }
            entracn = false;

            //---Notas de credito y cargo por ACONDICIONAMIENTO DE EMPAQUE EN DESTINO 10/05/2021
            can = 0;
            tot = 0;
            entracn = false;
            DataRow dtncr3;
            if (tablanc.Rows.Count > 0)
            {
                foreach (DataRow rwnc in tablanc.Select("lin_clave = '9815' and dnc_tipo = 'NCR'"))
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncr3 = tcon.NewRow();
                    dtncr3["cve_con"] = "109";
                    dtncr3["nombre_con"] = "Otros Conceptos Comisión";
                    dtncr3["unidades"] = can;
                    dtncr3["precio"] = Math.Round(tot / can, 3);
                    dtncr3["total"] = Math.Round(tot * -1, 3);
                    dtncr3["calculo"] = "1";
                    tcon.Rows.Add(dtncr3);

                    //var_dec_total = var_dec_total - Math.Round(tot, 3);
                }

            }
            entracn = false;

            can = 0;
            tot = 0;
            DataRow dtncg2;
            if (tablanc.Rows.Count > 0)
            {
                DataTable dtvw = new DataTable();
                DataView dw = tablanc.DefaultView;
                dw.RowFilter = "lin_clave = '9814' and dnc_tipo = 'NCG'";
                dtvw = dw.ToTable();

                foreach (DataRow rwnc in dtvw.Rows)
                {
                    can = can + Convert.ToDecimal(rwnc["dnc_cantidad"].ToString());
                    imp = Convert.ToDecimal(rwnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rwnc["dnc_precio_usd"].ToString());
                    tot = tot + imp;
                    imp = 0;
                    entracn = true;
                }
                if (entracn == true)
                {
                    dtncg2 = tcon.NewRow();
                    dtncg2["cve_con"] = "107";
                    dtncg2["nombre_con"] = "Notas de Cargo x Acond. Emp. Origen";
                    dtncg2["unidades"] = can;
                    dtncg2["precio"] = Math.Round(tot / can, 3);
                    dtncg2["total"] = Math.Round(tot, 3);//dtncg["total"] = Math.Round(tot * -1, 3);
                    dtncg2["calculo"] = "1";
                    tcon.Rows.Add(dtncg2);
                }

            }
            //---Fin Notas de credito y cargo por ACONDICIONAMIENTO DE EMPAQUE EN DESTINO

            //FIN MERMAS Y RECLAMACIONES, NOTAS DE CREDITO, NOTAS DE CARGO, RECHAZOS POR CALIDAD

            //PRESTAMOS
            //can = 0;
            //tot = 0;
            //foreach (DataRow rpre in dtPrestamos.Rows)
            //{
            //    bool fnd = lbl_producto.Text.Contains(rpre["Descripcion_Art"].ToString());
            //    if (fnd == true)
            //    {
            //        DataRow rwpre = tcon.NewRow();
            //        rwpre["cve_con"] = "95";
            //        rwpre["nombre_con"] = "DESCUENTO PAGO ANTICIPADO";
            //        rwpre["unidades"] = "1";
            //        decimal res = 0;
            //        res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
            //        rwpre["precio"] = Math.Round(res, 3);
            //        rwpre["total"] = Math.Round(res, 3) * -1;
            //        tcon.Rows.Add(rwpre);
            //        lblIdPrestamo.Text = rpre["Id_Movimiento"].ToString();
            //        break;
            //    }
            //}
            //FIN PRESTAMOS


            //AGREGAR CONCEPTO DE IMPORTE
            //can = 0;
            //tot = 0;
            //DataRow dtrimporte = tcon.NewRow();
            //dtrimporte["cve_con"] = "99";
            //dtrimporte["nombre_con"] = "Impuesto";
            //dtrimporte["unidades"] = Convert.ToDecimal(var_dec_unidades);
            //decimal cto_imp = 0;
            //if (lbl_producto.Text.Contains("SUNSET") == true)
            //{
            //    cto_imp = Convert.ToDecimal(var_dec_precio) * Convert.ToDecimal("0.175");
            //}
            //dtrimporte["precio"] = cto_imp;
            //dtrimporte["total"] = Math.Round((Convert.ToDecimal(cto_imp) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
            //tcon.Rows.Add(dtrimporte);

            //AGREGAR CONCEPTO DE ARANCELES
            can = 0;
            tot = 0;
            DataRow dtrimporte = tcon.NewRow();
            dtrimporte["cve_con"] = "113";
            dtrimporte["nombre_con"] = "COSTO ARANCELES";
            dtrimporte["unidades"] = Convert.ToDecimal(var_dec_unidades);
            decimal cto_imp = 0;
            if (lbl_producto.Text.Contains("TOMATE") == true)
            {
                SqlDataAdapter adapSql;
                adapSql = new SqlDataAdapter("spSISEMPPorcentajeImpueto", thisConnection);
                adapSql.SelectCommand.CommandType = CommandType.StoredProcedure;
                DataTable dtPorcentaje = new DataTable();
                adapSql.Fill(dtPorcentaje);
                adapSql.Dispose();

                decimal porcentaje = Convert.ToDecimal(dtPorcentaje.Rows[0]["Porcentaje"].ToString());

                SqlCommand COMMAND = new SqlCommand();
                COMMAND = thisConnection.CreateCommand();
                COMMAND.CommandText = ";WITH TotalUnidadesPorFolio AS ( " +
                                    "SELECT " +
                                        "A.fcn_folio, " +
                                        "SUM(A.fcn_num_unidades) AS total_unidades " +
                                    "FROM " +
                                        "tb_det_facturas A " +
                                        "INNER JOIN tb_mstr_facturas_nal B ON A.fcn_folio = B.fcn_folio AND A.fcn_tipo = B.fcn_lugar " +
                                    "where " +
                                        "B.fcn_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                                        "AND B.fcn_estatus <> 'C' " +
                                        "AND B.um_clave = 'USD' " +
                                        "AND B.fcn_monto <> B.ncr_monto " +
                                        "GROUP BY " +
                                            "A.fcn_folio " +
                                ") " +
                                "SELECT " +
                                    "b.prod_clave, " +
                                    "SUM(b.fcn_num_unidades) AS cajas, " +
                                    "SUM( " +
                                    "(b.fcn_precio_usd - ROUND(ISNULL(a.fcn_monto_transporte / NULLIF(t.total_unidades, 0), 0), 2)) " +
                                        "* b.fcn_num_unidades " +
                                    ") AS importe_ajustado " +
                                "FROM " +
                                    "tb_mstr_facturas_nal a " +
                                "JOIN " +
                                    "tb_det_facturas b ON a.fcn_folio = b.fcn_folio AND b.fcn_tipo = a.fcn_lugar " +
                                "JOIN " +
                                    "TotalUnidadesPorFolio t ON t.fcn_folio = a.fcn_folio " +
                                "WHERE " +
                                    "a.fcn_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                                    "AND a.fcn_estatus <> 'C' " +
                                    "AND a.um_clave = 'USD' " +
                                    "AND b.prod_clave = '" + this.cveprod + "' " +
                                    "AND a.fcn_monto <> a.ncr_monto " +
                                "GROUP BY " +
                                    "b.prod_clave " +
                                "ORDER BY " +
                                    "b.prod_clave;";
                SqlDataReader READERARANCEL;
                READERARANCEL = COMMAND.ExecuteReader();
                Decimal caixas = 0;
                Decimal import = 0;
                Decimal costoa = 0;
                if (READERARANCEL.HasRows)
                {
                    READERARANCEL.Read();
                    caixas = Convert.ToDecimal(READERARANCEL["cajas"].ToString());
                    import = Convert.ToDecimal(READERARANCEL["importe_ajustado"].ToString());
                }
                if (import == 0 && caixas == 0)
                    costoa = 0;
                else
                    costoa = Math.Round((import / caixas), 2);
                //cto_imp = Convert.ToDecimal(var_dec_precio) * Convert.ToDecimal("0.1709");
                cto_imp = Convert.ToDecimal(costoa) * porcentaje;//Convert.ToDecimal("0.169752");
            }
            dtrimporte["precio"] = cto_imp;
            dtrimporte["total"] = Math.Round((Convert.ToDecimal(cto_imp) * Convert.ToDecimal(var_dec_unidades)), 3) * -1;
            tcon.Rows.Add(dtrimporte);

            //SERVICIO DE LOGISTICA

            //if (lbl_cveprod.Text == "16TORAPO11")
            //{
            //    tablanc.Clear();
            //    cmnd1 = thisConnection.CreateCommand();
            //    //string mprov = "";
            //    mprov = lbl_cveprov.Text;
            //    cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
            //        "WHERE  A.clavep = 'SERIVICIOL' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '99' " +
            //        "AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
            //        "AND B.nc_folio = '1302' ORDER BY cveprov, prod_clave, fechap";
            //    DataRow drLOG;
            //    reader1 = cmnd1.ExecuteReader();
            //    while (reader1.Read())
            //    {
            //        drLOG = tablanc.NewRow();
            //        drLOG["prod_nombre"] = reader1.GetValue(0).ToString().Trim();
            //        drLOG["nc_folio"] = reader1.GetValue(1).ToString().Trim();
            //        drLOG["dnc_cantidad"] = reader1.GetValue(2).ToString().Trim();
            //        drLOG["dnc_precio_mn"] = reader1.GetValue(3).ToString().Trim();
            //        drLOG["dnc_precio_usd"] = reader1.GetValue(4).ToString().Trim();
            //        drLOG["clavep"] = reader1.GetValue(5).ToString().Trim();
            //        drLOG["dnc_tipo"] = reader1.GetValue(6).ToString().Trim();
            //        drLOG["lin_clave"] = reader1.GetValue(7).ToString().Trim();
            //        tablanc.Rows.Add(drLOG);
            //    }
            //    reader1.Close();
            //    reader1.Dispose();
            //    cmnd1.Dispose();

            //    decimal impLog = 0;
            //    can = 0;
            //    tot = 0;
            //    DataRow dtLOG;
            //    if (tablanc.Rows.Count > 0)
            //    {
            //        bool ent = false;
            //        foreach (DataRow rnc in tablanc.Rows)
            //        {
            //            ent = true;
            //            can = can + Convert.ToDecimal(rnc["dnc_cantidad"].ToString());
            //            impLog = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
            //            //imp = can * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
            //            tot = tot + impLog;
            //            impLog = 0;
            //        }

            //        if (ent == true)
            //        {
            //            dtLOG = tcon.NewRow();
            //            dtLOG["cve_con"] = "100";
            //            dtLOG["nombre_con"] = "Servicio de logistica";
            //            dtLOG["unidades"] = can;
            //            dtLOG["precio"] = Math.Round((tot / can), 3);
            //            dtLOG["total"] = tot;
            //            tcon.Rows.Add(dtLOG);
            //        }
            //    }
            //}
            //if (lbl_cveprod.Text == "16TOSURA11")
            //{
            tablanc.Clear();
            cmnd1 = thisConnection.CreateCommand();
            //string mprov = "";
            mprov = lbl_cveprov.Text;

            decimal cantidade = 0;

            cmnd1.CommandText = "SELECT SUM(A.dnc_cantidad) AS dnc_cantidad FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                "WHERE A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9812' " +
                "AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.clavep = '" + lbl_cveprod.Text + "' " +
                "AND A.liq_folio_exp = '0'";

            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                string c = reader1["dnc_cantidad"].ToString().Trim();
                if (reader1["dnc_cantidad"] == null || reader1["dnc_cantidad"].ToString().Trim() == "" || reader1["dnc_cantidad"].ToString().Trim() == "0")
                    cantidade = 0;
                else
                    cantidade = Convert.ToDecimal(reader1["dnc_cantidad"].ToString().Trim());
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            //cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
            //    "WHERE  A.clavep = 'SERIVICIOL' AND A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '99' " +
            //    "AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
            //    "AND B.nc_folio = '1314' ORDER BY cveprov, prod_clave, fechap";
            if (cantidade > 0)
            {
                cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap " +
                    "FROM tb_det_notascyc A " +
                    "INNER JOIN tb_mstr_notascyc B ON A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
                    "WHERE  (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9812' " +
                    "AND A.dnc_tipo = 'NCG' and B.nc_estatus <> 'C' AND A.clavep = '" + lbl_cveprod.Text + "' AND A.liq_folio_exp = '0' " +
                    "ORDER BY cveprov, prod_clave, fechap";
                DataRow drLOG;
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    drLOG = tablanc.NewRow();
                    drLOG["prod_nombre"] = reader1.GetValue(0).ToString().Trim();
                    drLOG["nc_folio"] = reader1.GetValue(1).ToString().Trim();
                    drLOG["dnc_cantidad"] = reader1.GetValue(2).ToString().Trim();// cantidade.ToString();
                    drLOG["dnc_precio_mn"] = reader1.GetValue(3).ToString().Trim();
                    drLOG["dnc_precio_usd"] = reader1.GetValue(4).ToString().Trim();
                    drLOG["clavep"] = reader1.GetValue(5).ToString().Trim();
                    drLOG["dnc_tipo"] = reader1.GetValue(6).ToString().Trim();
                    drLOG["lin_clave"] = reader1.GetValue(7).ToString().Trim();
                    tablanc.Rows.Add(drLOG);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                decimal impLog = 0;
                can = 0;
                tot = 0;
                DataRow dtLOG;
                if (tablanc.Rows.Count > 0)
                {
                    bool ent = false;
                    foreach (DataRow rnc in tablanc.Rows)
                    {
                        ent = true;
                        can = can + Convert.ToDecimal(rnc["dnc_cantidad"].ToString());
                        impLog = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                        //imp = can * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                        tot = tot + impLog;
                        impLog = 0;
                    }

                    if (ent == true)
                    {
                        dtLOG = tcon.NewRow();
                        dtLOG["cve_con"] = "100";
                        dtLOG["nombre_con"] = "Servicio de logistica";
                        dtLOG["unidades"] = can;
                        dtLOG["precio"] = Math.Round((tot / can), 3);
                        dtLOG["total"] = tot;
                        dtLOG["calculo"] = "1";
                        tcon.Rows.Add(dtLOG);
                    }
                }
            }

            //}
            //FIN SERVICIO DE LOGISTICA

            //----------13/03/2024----------//
            //SERVICIO DE LOGISTICA USDA SOLICITADO POR LISSETE
            tablanc.Clear();
            cmnd1 = thisConnection.CreateCommand();
            mprov = lbl_cveprov.Text;
            decimal cantidade2 = 0;

            cmnd1.CommandText = "SELECT SUM(A.dnc_cantidad) AS dnc_cantidad FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                "WHERE A.fechap BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "' AND A.lin_clave = '9812' AND A.prod_clave = '981218' " +
                "AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.clavep = '" + lbl_cveprod.Text + "' AND A.liq_folio_exp = '0'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                string c = reader1["dnc_cantidad"].ToString().Trim();
                if (reader1["dnc_cantidad"] == null || reader1["dnc_cantidad"].ToString().Trim() == "" || reader1["dnc_cantidad"].ToString().Trim() == "0")
                    cantidade2 = 0;
                else
                    cantidade2 = Convert.ToDecimal(reader1["dnc_cantidad"].ToString().Trim());
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            if (cantidade2 > 0)
            {
                cmnd1.CommandText = "SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap " +
                    "FROM tb_det_notascyc A " +
                    "INNER JOIN tb_mstr_notascyc B ON A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
                    "WHERE  (A.fechap >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND A.fechap <= '" + Convert.ToDateTime(f2).ToShortDateString() + "') AND A.lin_clave = '9812' " +
                    "AND A.prod_clave = '981218' AND A.dnc_tipo = 'NCR' and B.nc_estatus <> 'C' AND A.clavep = '" + lbl_cveprod.Text + "' AND A.liq_folio_exp = '0' " +
                    "ORDER BY cveprov, prod_clave, fechap";
                DataRow drLOG;
                reader1 = cmnd1.ExecuteReader();
                while (reader1.Read())
                {
                    drLOG = tablanc.NewRow();
                    drLOG["prod_nombre"] = reader1.GetValue(0).ToString().Trim();
                    drLOG["nc_folio"] = reader1.GetValue(1).ToString().Trim();
                    drLOG["dnc_cantidad"] = reader1.GetValue(2).ToString().Trim();// cantidade.ToString();
                    drLOG["dnc_precio_mn"] = reader1.GetValue(3).ToString().Trim();
                    drLOG["dnc_precio_usd"] = reader1.GetValue(4).ToString().Trim();
                    drLOG["clavep"] = reader1.GetValue(5).ToString().Trim();
                    drLOG["dnc_tipo"] = reader1.GetValue(6).ToString().Trim();
                    drLOG["lin_clave"] = reader1.GetValue(7).ToString().Trim();
                    tablanc.Rows.Add(drLOG);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();

                decimal impLog = 0;
                can = 0;
                tot = 0;
                DataRow dtLOG;
                if (tablanc.Rows.Count > 0)
                {
                    bool ent = false;
                    foreach (DataRow rnc in tablanc.Rows)
                    {
                        ent = true;
                        can = can + Convert.ToDecimal(rnc["dnc_cantidad"].ToString());
                        impLog = Convert.ToDecimal(rnc["dnc_cantidad"].ToString()) * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                        //imp = can * Convert.ToDecimal(rnc["dnc_precio_usd"].ToString());
                        tot = tot + impLog;
                        impLog = 0;
                    }

                    if (ent == true)
                    {
                        dtLOG = tcon.NewRow();
                        dtLOG["cve_con"] = "110";
                        dtLOG["nombre_con"] = "Servicio de logistica USDA";
                        dtLOG["unidades"] = can;
                        dtLOG["precio"] = Math.Round((tot / can), 3);
                        dtLOG["total"] = Math.Round(tot * -1, 3);
                        dtLOG["calculo"] = "1";
                        tcon.Rows.Add(dtLOG);
                    }
                }
            }



            //FIN SERVICIO DE LOGISTICA
            //----------FIN 13/03/2024----------//

            //PRESTAMOS
            can = 0;
            tot = 0;
            //PRESTAMOS
            //----------29/11/2017----------//
            foreach (DataRow rpre in dtPrestamos.Rows)
            {
                DataRow rwpre = tcon.NewRow();
                rwpre["cve_con"] = rpre["Lin_Clave"].ToString();
                rwpre["nombre_con"] = rpre["Descripcion_Art"].ToString() + " REF: " + rpre["Factura"].ToString();
                rwpre["unidades"] = "1";
                decimal res = 0;
                decimal res_moneda = 0;
                res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
                if (txt_tipo.Text == "NACIONAL")//NACIONAL
                {
                    if (rpre["Moneda"].ToString() == "DOLARES")
                        res_moneda = res * Convert.ToDecimal(lblTipoCambio.Text);
                    else
                        res_moneda = res;
                }
                if (txt_tipo.Text == "EXPORTACION")//EXPORTACION
                {
                    if (rpre["Moneda"].ToString() == "PESOS")
                        res_moneda = res / Convert.ToDecimal(lblTipoCambio.Text);
                    else
                        res_moneda = res;
                }
                rwpre["precio"] = Math.Round(res_moneda, 3);
                rwpre["total"] = Math.Round(res_moneda, 3) * -1;
                rwpre["valor"] = rpre["Id_Movimiento"].ToString();
                rwpre["moni"] = rpre["Moneda"].ToString();
                rwpre["saldo"] = Math.Round(res_moneda, 3);
                tcon.Rows.Add(rwpre);
            }
            //----------FIN 29/11/2017----------//
            //foreach (DataRow rpre in dtPrestamos.Rows)
            //{
            //    DataRow rwpre = tcon.NewRow();
            //    rwpre["cve_con"] = rpre["Lin_Clave"].ToString();
            //    rwpre["nombre_con"] = rpre["Descripcion_Art"].ToString();
            //    rwpre["unidades"] = "1";
            //    decimal res = 0;
            //    res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
            //    rwpre["precio"] = Math.Round(res, 3);
            //    rwpre["total"] = Math.Round(res, 3) * -1;
            //    rwpre["valor"] = rpre["Id_Movimiento"].ToString();
            //    tcon.Rows.Add(rwpre);

            //    //bool fnd = lbl_producto.Text.Contains(rpre["Descripcion_Art"].ToString());
            //    //if (fnd == true)
            //    //{
            //    //DataRow rwpre = tcon.NewRow();
            //    //rwpre["cve_con"] = "95";
            //    //rwpre["nombre_con"] = "DESCUENTO PAGO ANTICIPADO";
            //    //rwpre["unidades"] = "1";
            //    //decimal res = 0;
            //    //res = Convert.ToDecimal(rpre["Total"]) - Convert.ToDecimal(rpre["Saldo"]);
            //    //rwpre["precio"] = Math.Round(res, 3);
            //    //rwpre["total"] = Math.Round(res, 3) * -1;
            //    //tcon.Rows.Add(rwpre);
            //    //lblIdPrestamo.Text = rpre["Id_Movimiento"].ToString();
            //    //    break;
            //    //}
            //}
            //FIN PRESTAMOS


            //FIN PRESTAMOS

            //----------CALCULO DE MERMAS GAB----------
            //can = 0;
            //tot = 0;
            //DataTable dtMermasGab = new DataTable();
            //SqlDataAdapter adap = new SqlDataAdapter("spSISEMPLiquidacionesMermaGab", thisConnection);
            //adap.SelectCommand.CommandType = CommandType.StoredProcedure;
            //adap.SelectCommand.Parameters.AddWithValue("@FechaI", Convert.ToDateTime(f1).ToShortDateString());
            //adap.SelectCommand.Parameters.AddWithValue("@FechaF", Convert.ToDateTime(f2).ToShortDateString());
            //adap.SelectCommand.Parameters.AddWithValue("@Prod", cveprod);
            //adap.Fill(dtMermasGab);

            //decimal TCj = 0;
            //decimal Flt = 0;
            //decimal MyR = 0;
            //decimal Com = 0;
            //decimal TCo = 0;

            //tcon.AsEnumerable()
            //    .Where(row => row.Field<string>("cve_con") == "1")
            //    .ToList()
            //    .ForEach(row =>
            //    {
            //        TCj = Convert.ToDecimal(row["unidades"]);
            //    });

            //tcon.AsEnumerable()
            //    .Where(row => row.Field<string>("cve_con") == "6")
            //    .ToList()
            //    .ForEach(row =>
            //    {
            //        Flt = Math.Round((Convert.ToDecimal(row["total"]) / Convert.ToDecimal(row["unidades"]) * -1), 3);
            //    });

            //tcon.AsEnumerable()
            //    .Where(row => row.Field<string>("cve_con") == "7")
            //    .ToList()
            //    .ForEach(row =>
            //    {
            //        MyR = Math.Round((Convert.ToDecimal(row["total"]) * -1) / TCj, 3);
            //    });


            //Com = Comision();

            //TCo = Math.Round((Com * -1) / TCj, 3);

            //Decimal tot_dec_merma = Flt + MyR + TCo;

            //foreach (DataRow rt in dtMermasGab.Rows)
            //{
            //    decimal cto_mem = Convert.ToDecimal(rt["Costo"].ToString()) - tot_dec_merma;
            //    tot = Convert.ToDecimal(rt["Cantidad"].ToString()) * cto_mem;
            //    can = Convert.ToDecimal(rt["Cantidad"].ToString());
            //    dtncg2 = tcon.NewRow();
            //    dtncg2["cve_con"] = "111";
            //    dtncg2["nombre_con"] = "Merma en Planta";
            //    dtncg2["unidades"] = can;
            //    dtncg2["precio"] = Math.Round(cto_mem, 3);//Math.Round(tot / can, 3);
            //    dtncg2["total"] = Math.Round(tot * -1, 3);//dtncg["total"] = Math.Round(tot * -1, 3);
            //    dtncg2["calculo"] = "1";
            //    tcon.Rows.Add(dtncg2);
            //}
            //----------FIN CALCULO DE MERMAS GAB----------

            thisConnection.Close();

            for (int i = 0; i < tcon.Rows.Count; i++)
            {
                dtgConceptos.Rows.Add(tcon.Rows[i]["cve_con"].ToString(), tcon.Rows[i]["nombre_con"].ToString(), Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()).ToString("###,###,##0.000"),
                    Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("###,###,##0.000"), tcon.Rows[i]["valor"].ToString(), tcon.Rows[i]["moni"].ToString(), "", tcon.Rows[i]["saldo"].ToString(), tcon.Rows[i]["calculo"].ToString());
            }

            thisConnection.Close();

            foreach (DataGridViewRow gr in dtgConceptos.Rows)
            {
                if (gr.Cells["clave"].Value.ToString().Length > 4)
                {
                    if (Convert.ToDecimal(gr.Cells["precio"].Value.ToString()) == 0)
                    {
                        gr.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
        }

        private void historicoempaque()
        {
            thisConnection.Open();

            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "SELECT emp_clave, (sum(hrp_cantidad * hrp_costo) / sum(hrp_cantidad)) as hrp_costo FROM tb_historico_recepcion" +
                " WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'" +
                " AND hrp_costo > 0 AND hrp_estatus <> 'C' AND (hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                " AND alm_clave in ('01', '02') " +// + query +
                " GROUP BY hrp_tipo_recepcion, emp_clave" +
                " ORDER BY hrp_tipo_recepcion, emp_clave";
            reader2 = cmnd2.ExecuteReader();
            DataRow tpperow;
            while (reader2.Read())
            {
                tpperow = tppe.NewRow();
                tpperow["emp_clave"] = reader2.GetValue(0).ToString().Trim();
                tpperow["hrp_costo"] = reader2.GetDecimal(1);
                tppe.Rows.Add(tpperow);
            }
            reader2.Close();
            reader2.Dispose();
            cmnd2.Dispose();

            cmnd2.CommandText = "SELECT emp_clave, emp_costo FROM tb_cat_empaques WHERE alm_clave in ('01', '02')  ORDER BY emp_clave";//" + query + "
            reader2 = cmnd2.ExecuteReader();
            DataRow ectorow;
            while (reader2.Read())
            {
                ectorow = ecto.NewRow();
                ectorow["emp_clave"] = reader2.GetValue(0).ToString().Trim();
                ectorow["emp_costo"] = reader2.GetDecimal(1);
                ecto.Rows.Add(ectorow);
            }
            reader2.Close();
            reader2.Dispose();
            cmnd2.Dispose();

            thisConnection.Close();
        }

        private DataSet historicoempaque(string emp)
        {
            DataSet ds = new DataSet();

            //thisConnection.Open();

            SqlDataAdapter ad = new SqlDataAdapter("SELECT emp_clave, (sum(hrp_cantidad * hrp_costo) / sum(hrp_cantidad)) as hrp_costo FROM tb_historico_recepcion" +
                " WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'" +
                " AND hrp_costo > 0 AND hrp_estatus <> 'C' AND (hrp_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                " AND alm_clave = '02' AND emp_clave = '" + emp + "' " +// + query +
                " GROUP BY hrp_tipo_recepcion, emp_clave" +
                " ORDER BY hrp_tipo_recepcion, emp_clave", thisConnection);
            ad.Fill(ds, "historico");

            ad = new SqlDataAdapter("SELECT emp_clave, emp_costo FROM tb_cat_empaques WHERE alm_clave = '02' AND emp_clave = '" + emp + "'  ORDER BY emp_clave", thisConnection);
            ad.Fill(ds, "catalogo");

            //thisConnection.Close();

            return ds;
        }

        private decimal precioempaque(DateTime fech2)
        {
            decimal tot1x = 0;
            decimal tot2x = 0;
            decimal tot3x = 0;
            //List<string> listaempaques = new List<string>();
            //thisConnection.Open();
            cmnd3 = thisConnection.CreateCommand();
            string fe1 = "";
            cmnd3.CommandText = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp ORDER BY invemp_fecha DESC";
            reader3 = cmnd3.ExecuteReader();
            if (reader3.HasRows)
            {
                reader3.Read();
                fe1 = Convert.ToDateTime(reader3.GetValue(0).ToString().Trim()).ToString("dd-MM-yyyy");//ULTIMO CIERRE
            }
            reader3.Close();
            reader3.Dispose();
            cmnd3.Dispose();

            cmnd3 = thisConnection.CreateCommand();
            //cmnd5.CommandText = "SELECT emp_clave, AVG(hrp_costo) as hrp_costo FROM tb_historico_recepcion" +
            //    " WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'" +
            //    " AND hrp_costo > 0 AND hrp_estatus <> 'C' AND (hrp_fecha >= '" + this.f1.ToShortDateString() + "' AND hrp_fecha <= '" + this.f2.ToShortDateString() + "')" +
            //    " AND alm_clave in ('01', '02') " +// + query +
            //    " GROUP BY hrp_tipo_recepcion, emp_clave" +
            //    " ORDER BY hrp_tipo_recepcion, emp_clave";
            cmnd3.CommandText = "SELECT emp_clave, SUM(hrp_cantidad * hrp_costo) as total FROM tb_historico_recepcion" +
                " WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'" +
                " AND hrp_estatus <> 'C' AND (hrp_fecha >= '" + fe1 + "' AND hrp_fecha <= '" + fech2.ToShortDateString() + "')" +
                " AND alm_clave in ('01', '02') AND emp_clave in ('M2628', 'N3742')" +// + query +
                " GROUP BY hrp_tipo_recepcion, emp_clave" +
                " ORDER BY hrp_tipo_recepcion, emp_clave";
            reader3 = cmnd3.ExecuteReader();
            while (reader3.Read())
            {
                tot1x = tot1x + Convert.ToDecimal(reader3.GetValue(1).ToString().Trim());
            }
            reader3.Close();
            reader3.Dispose();
            cmnd3.Dispose();

            cmnd3 = thisConnection.CreateCommand();
            cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades) AS exportacion FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + fe1 + "' and F.fcn_fecha <= '" + fech2.ToShortDateString() + "' " +
                "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%APIO%' AND DF.lin_clave in ('01', '16') AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                "GROUP BY DF.lin_clave, DF.prod_clave " +
                "ORDER BY DF.lin_clave, DF.prod_clave";
            reader3 = cmnd3.ExecuteReader();
            while (reader3.Read())
            {
                tot2x = tot2x + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
            }

            tot3x = tot1x / tot2x;

            //cmnd5.CommandText = "SELECT emp_clave, emp_costo FROM tb_cat_empaques WHERE alm_clave in ('01', '02')  ORDER BY emp_clave";//" + query + "
            //reader5 = cmnd5.ExecuteReader();
            //DataRow ectorow;
            //while (reader5.Read())
            //{
            //    ectorow = ecto.NewRow();
            //    ectorow["emp_clave"] = reader5.GetValue(0).ToString().Trim();
            //    ectorow["emp_costo"] = reader5.GetDecimal(1);
            //    ecto.Rows.Add(ectorow);
            //}
            //reader5.Close();
            //reader5.Dispose();
            //cmnd5.Dispose();

            //thisConnection.Close();
            return tot3x;
        }

        private void calculatotales()
        {
            decimal subtotal = 0;
            decimal porcen = 0;
            decimal descuento = 0;

            decimal subtotal_1 = 0;
            foreach (DataRow rw in tcon.Rows)
            {
                //if(rw["cve_con"].ToString() != "93")
                //    subtotal = subtotal + (Convert.ToDecimal(rw["total"].ToString()));
                //else
                subtotal = subtotal + (Convert.ToDecimal(rw["total"].ToString()));

                if (rw["cve_con"].ToString() == "93" || rw["cve_con"].ToString() == "95" || rw["cve_con"].ToString() == "100" || rw["cve_con"].ToString() == "102"
                     || rw["cve_con"].ToString() == "103" || rw["cve_con"].ToString() == "104" || rw["cve_con"].ToString() == "105" || rw["cve_con"].ToString() == "106")
                    subtotal_1 += Convert.ToDecimal(rw["total"].ToString());

            }

            subtotal_1 = subtotal_1 + Convert.ToDecimal(tcon.Rows[0]["total"].ToString());

            if (txt_valor_por.Text != "")
            {
                //porcen = Convert.ToDecimal(tcon.Rows[0]["total"].ToString()) * (Convert.ToDecimal(txt_valor_por.Text) / 100) * -1;
                porcen = Convert.ToDecimal(subtotal_1) * (Convert.ToDecimal(txt_valor_por.Text) / 100) * -1;
            }
            txt_porcentaje.Text = porcen.ToString("###,###,##0.000");

            if (txt_porce_desc.Text != "")
            {
                //descuento = Convert.ToDecimal(tcon.Rows[0]["total"].ToString()) * (Convert.ToDecimal(txt_porce_desc.Text) / 100) * -1;
                descuento = Convert.ToDecimal(subtotal_1) * (Convert.ToDecimal(txt_porce_desc.Text) / 100) * -1;
            }

            txt_total.Text = subtotal.ToString("###,###,##0.000");
            txt_cant_porce.Text = descuento.ToString("###,###,##0.000");
            txt_liquidar.Text = (subtotal + porcen + descuento).ToString("###,###,##0.000");
            txt_costounitario.Text = (Convert.ToDecimal(txt_liquidar.Text) / Convert.ToDecimal(tcon.Rows[0]["unidades"].ToString())).ToString("###,###,##0.000");



        }

        private decimal Comision()
        {
            decimal subtotal = 0;
            decimal porcen = 0;

            decimal subtotal_1 = 0;
            foreach (DataRow rw in tcon.Rows)
            {
                subtotal = subtotal + (Convert.ToDecimal(rw["total"].ToString()));

                if (rw["cve_con"].ToString() == "93" || rw["cve_con"].ToString() == "95" || rw["cve_con"].ToString() == "100" || rw["cve_con"].ToString() == "102"
                     || rw["cve_con"].ToString() == "103" || rw["cve_con"].ToString() == "104" || rw["cve_con"].ToString() == "105" || rw["cve_con"].ToString() == "106")
                    subtotal_1 += Convert.ToDecimal(rw["total"].ToString());

            }
            subtotal_1 = subtotal_1 + Convert.ToDecimal(tcon.Rows[0]["total"].ToString());

            if (txt_valor_por.Text != "")
            {
                porcen = Convert.ToDecimal(subtotal_1) * (Convert.ToDecimal(txt_valor_por.Text) / 100) * -1;
            }

            return porcen;

        }

        private void descuentoautoservicio()
        {
            //DESCOMENTAR
            if (cveprov == "01" || cveprov == "03" || cveprov == "1328")
            {
                if (procedencia == "NACIONAL")
                {
                    //NUEVO PROCESO PARA DESCUENTO DE AUTOSERVICIO 22/12/2015 
                    thisConnection.Open();
                    cmnd3 = thisConnection.CreateCommand();
                    cmnd3.CommandText = "select D.clavep, SUM(D.dnc_precio_mn) from tb_det_notascyc D, tb_mstr_notascyc M" +
                        " where D.nc_folio = M.nc_folio and D.dnc_tipo = M.nc_tipo AND D.dnc_tipo = 'NCR' and D.dnc_lugar = M.nc_lugar and M.nc_lugar <> 'EXP' AND D.prod_nombre = 'OTROS CONCEPTOS'" +
                        " AND M.nc_devbon = D.dnc_devbon and (M.nc_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and M.nc_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                        " GROUP BY D.clavep" +
                        " order by D.clavep";
                    reader3 = cmnd3.ExecuteReader();
                    decimal totnotascr = 0;
                    if (reader3.HasRows)
                    {
                        while (reader3.Read())
                        {
                            totnotascr = totnotascr + Convert.ToDecimal(reader3.GetValue(1).ToString().Trim());
                        }
                    }
                    reader3.Close();
                    reader3.Dispose();
                    cmnd3.Dispose();
                    cmnd3 = thisConnection.CreateCommand();
                    cmnd3.CommandText = "SELECT DF.lin_clave, DF.prod_clave, SUM(DF.fcn_num_unidades * DF.fcn_precio_mn) AS nacional, P.prod_nombre " +
                        "FROM tb_det_facturas DF, tb_mstr_facturas_nal F, tb_cat_producto p " +
                        "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + Convert.ToDateTime(f1).ToShortDateString() + "' and F.fcn_fecha <= '" + Convert.ToDateTime(f2).ToShortDateString() + "' " +
                        "AND F.fcn_lugar <> 'EXP' AND DF.prod_clave = P.prod_clave " +
                        "AND DF.fcn_tipo = F.fcn_lugar AND F.fcn_monto <> F.ncr_monto " +
                        "GROUP BY DF.lin_clave, DF.prod_clave, P.prod_nombre ORDER BY DF.lin_clave, DF.prod_clave";
                    reader3 = cmnd3.ExecuteReader();
                    decimal totventasnal = 0;
                    if (reader3.HasRows)
                    {
                        while (reader3.Read())
                        {
                            totventasnal = totventasnal + Convert.ToDecimal(reader3.GetValue(2).ToString().Trim());
                        }
                    }
                    reader3.Close();
                    reader3.Dispose();
                    cmnd3.Dispose();
                    thisConnection.Close();
                    decimal porcentaje = 0;
                    porcentaje = (totnotascr / totventasnal) * 100;
                    txt_porce_desc.Text = porcentaje.ToString("###,##0.000");
                    KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                    txt_porce_desc_KeyPress(null, llave);
                    //FIN NUEVO PROCESO

                    calculatotales();
                }
            }
        }

        private void txt_porce_desc_KeyPress(object sender, KeyPressEventArgs e)
        {
            decimal totdesc = 0;
            if (e.KeyChar == Convert.ToChar(Keys.Enter))
            {
                if (txt_porce_desc.Text == "")
                {
                    MessageBox.Show("Debe ingresar un porcentaje", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                totdesc = Convert.ToDecimal(tcon.Rows[0]["total"].ToString()) * (Convert.ToDecimal(txt_porce_desc.Text) / 100) * -1;
                txt_cant_porce.Text = totdesc.ToString("###,###,##0.000");
            }
        }

        private void txt_valor_por_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                calculoporcentaje();
            }
        }

        private void calculoporcentaje()
        {

            decimal por = (txt_valor_por.Text == "") ? 0 : Convert.ToDecimal(txt_valor_por.Text);
            string val_cant = dtgConceptos.Rows[0].Cells[4].Value.ToString();
            decimal cant = Convert.ToDecimal(val_cant);

            lbl_porcentaje.Text = "% " + por.ToString();

            decimal p = 0;

            p = (cant * ((por) / 100));
            txt_porcentaje.Text = Convert.ToDecimal(p * -1).ToString("###,###,##0.000");



            //decimal t = 0;
            //t = Convert.ToDecimal(txt_total.Text) - (p);

            calculatotales();

        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public bool validarnumero(string dato)
        {
            try
            {
                decimal num = Convert.ToDecimal(dato);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void txt_tipocambio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {

            }
        }

        private void txt_precio_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (DDLDescuento.SelectedIndex == -1)
                {
                    MessageBox.Show("Debe seleccionar un concepto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (txt_unidades.Text == "")
                {
                    MessageBox.Show("Debe ingresar las unidades", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (txt_precio.Text == "")
                {
                    MessageBox.Show("Debe ingresar las unidades", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (validarnumero(txt_precio.Text) == false)
                {
                    MessageBox.Show("El valor ingresado no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (Convert.ToDecimal(txt_precio.Text) <= 0)
                {
                    MessageBox.Show("El precio debe ser mayor a cero", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string cve_desc = "";
                string nom_desc = "";
                decimal uni_desc = 0;
                decimal pre_desc = 0;

                cve_desc = clavedesc[DDLDescuento.SelectedIndex].ToString();
                nom_desc = DDLDescuento.SelectedItem.ToString();
                uni_desc = Convert.ToDecimal(txt_unidades.Text);
                pre_desc = Convert.ToDecimal(txt_precio.Text);


                if (txt_tipocambio.Text == "")
                {
                    if (cve_desc == "93" || cve_desc == "95" || cve_desc == "100" || cve_desc == "102" || cve_desc == "103" || cve_desc == "104" || cve_desc == "105" || cve_desc == "106")
                    {
                        dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), (uni_desc * pre_desc).ToString("###,###,##0.000"), "");
                        tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, (uni_desc * pre_desc), "");
                    }
                    else
                    {
                        dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), ((uni_desc * pre_desc) * -1).ToString("###,###,##0.000"), "");
                        tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, ((uni_desc * pre_desc) * -1), "");
                    }
                }
                else
                {
                    if (txtTipoCambioResp.Text == "")
                    {
                        if (cve_desc == "93" || cve_desc == "95" || cve_desc == "100" || cve_desc == "102" || cve_desc == "103" || cve_desc == "104" || cve_desc == "105" || cve_desc == "106")
                        {
                            dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), (uni_desc * pre_desc).ToString("###,###,##0.000"), "");
                            tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, (uni_desc * pre_desc), "");
                        }
                        else
                        {
                            dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), ((uni_desc * pre_desc) * -1).ToString("###,###,##0.000"), "");
                            tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, ((uni_desc * pre_desc) * -1), "");
                        }

                    }
                    else
                    {
                        if (cve_desc == "93" || cve_desc == "95" || cve_desc == "100" || cve_desc == "102" || cve_desc == "103" || cve_desc == "104" || cve_desc == "105" || cve_desc == "106")
                        {
                            //pre_desc = pre_desc / Convert.ToDecimal(txt_tipocambio.Text);
                            dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), (uni_desc * pre_desc).ToString("###,###,##0.000"), "");
                            tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, (uni_desc * pre_desc), "");
                        }
                        else
                        {
                            pre_desc = pre_desc / Convert.ToDecimal(txt_tipocambio.Text);
                            dtgConceptos.Rows.Add(cve_desc, nom_desc, Convert.ToDecimal(uni_desc).ToString("###,###,##0.000"), Convert.ToDecimal(pre_desc).ToString("###,###,##0.000"), ((uni_desc * pre_desc) * -1).ToString("###,###,##0.000"), "");
                            tcon.Rows.Add(cve_desc, nom_desc, uni_desc, pre_desc, ((uni_desc * pre_desc) * -1), "");
                        }

                    }

                }

                calculatotales();
            }
        }

        private void btnGuarda_Click(object sender, EventArgs e)
        {
            if (txtTipoLiq.Text == "nueva")
            {
                if (txt_tipo.Text == "NACIONAL")
                {
                    guardanuevonal();
                }
                if (txt_tipo.Text == "EXPORTACION")
                {
                    guardanuevoexp();
                }
            }
            if (txtTipoLiq.Text == "nuevotipo")
            {
                if (txt_tipo.Text == "NACIONAL")
                {
                    guardafaltantenal();
                }
                if (txt_tipo.Text == "EXPORTACION")
                {
                    guardafaltanteexp();
                }
            }
            if (txtTipoLiq.Text == "consulta")
            {
                if (txt_tipo.Text == "NACIONAL")
                {
                    modificarnacional();
                }
                if (txt_tipo.Text == "EXPORTACION")
                {
                    modificarexportacion();
                }
            }
        }

        private void guardanuevonal()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;
            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");
                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }





            string cvep = "";
            string var_chr_prod_clave = "";
            decimal canti = 0;

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;
            canti = 0;

            if (var_chr_prod_clave == "05005LETOR" || var_chr_prod_clave == "05005LETOT")
            {
                if (tcon.Rows[0][0].ToString() == "1")
                {
                    canti = Math.Round(Convert.ToDecimal(lbl_libras.Text) * Convert.ToDecimal(tcon.Rows[0]["unidades"].ToString()), 2);
                }
            }

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT COUNT(liq_folio) AS conteo FROM tb_mstr_liquidacion WHERE liq_fecha1 = '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                "AND liq_fecha2 = '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND liq_provcve = '" + lbl_cveprov.Text + "' " +
                "AND liq_lincve = '" + txt_lincve.Text + "' AND liq_prodcve = '" + lbl_cveprod.Text + "' AND uni_nac = '" + tcon.Rows[0]["unidades"].ToString() + "' AND status = 'A'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                Int32 val = Convert.ToInt32(reader1["conteo"].ToString());
                if (val > 0)
                {
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    return;
                }

            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();



                query = "INSERT INTO tb_mstr_liquidacion" +
                    "(liq_fecha1, liq_fecha2, liq_provcve, liq_lincve," +//
                    " liq_linnom, liq_provnom, liq_prodcve, liq_prodnom, liq_cajas_pal," +//
                    " liq_flejes_pal, liq_unidades, liq_mermas, uni_nac, uninac_oc," +//
                    " uni_mern, uni_exp, uniexp_oc, uni_mere, liq_pre_uni, " +//
                    " liq_porcen1, liq_porcen2, liq_imp_tot, liq_imp_por, liq_imp_liq," +//
                    " liq_costo1, liq_costo2, liq_afecto, liq_precambio, liq_preunie," +//
                    " liq_porcen_1e, liq_porcen_2e, liq_imp_tote, liq_imp_pore, liq_imp_liqe," +//
                    " liq_costo1e, liq_costo2e, status, liq_ocompra, liq_numoc," +//
                    " liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8," +
                    " liq_cantiocn, liq_cantioce, liq_libras, liq_por_des, liq_imp_pordes," +
                    " liq_exp_pordes, liq_exp_imppordes, tipo, liq_fecha, liq_nac, conse)" +
                    " VALUES" +
                    " ('" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "', '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "', '" + lbl_cveprov.Text + "', '" + txt_lincve.Text + "'," +
                    " '" + txt_linnom.Text + "', '" + lbl_proveedor.Text + "', '" + lbl_cveprod.Text + "', '" + lbl_producto.Text.Replace("'", " ").ToString() + "', '" + lbl_cajas.Text + "'," +
                    " '" + lbl_flejes.Text + "', '" + cantidad + "', 0, '" + var_dec_unidades + "', 0," +
                    " 0, 0, 0, 0, '" + var_dec_precio + "'," +
                    " '" + var_dec_prod_comision + "', '" + var_dec_prod_comision2 + "', '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.0000") + "', '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.0000") + "'," +
                    " '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.0000") + "', '" + ((txt_nuevocosto.Text == "") ? txt_nuevocosto.Text : "0") + "', '0', 0, 0," +//
                    " 0, 0, 0, 0, 0," +//
                    " 0, 0, 'A', '', ''," +
                    " '', '', '', '', '', '', '', ''," +
                    " 0, 0, '" + canti + "', '" + txt_porce_desc.Text + "', '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.0000") + "'," +
                    " 0, 0, '" + tl + "', '" + DateTime.Now.ToShortDateString() + "', '" + lblTeorico.Text + "', '" + ((lblConse.Text == "") ? "0" : lblConse.Text) + "') SELECT SCOPE_IDENTITY()";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                if (lbl_liquidacion.Text != ultimo_folio)
                {
                    MessageBox.Show("El folio de la liquidacion ha cambiado por movimientos en la red, el numero de folio asignado es: " + ultimo_folio + "\nSe imprimirá la nueva liquidacion enseguida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    lbl_liquidacion.Text = ultimo_folio;
                    //printDocument1.Print();
                    thisConnection.Close();
                    //printDocument1.Print();
                    thisConnection.Open();
                }
                else
                {
                    thisConnection.Close();
                    //printDocument1.Print();
                    thisConnection.Open();
                }

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Inserción de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", lbl_liquidacion.Text, "INSERCION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();

                    calc = tcon.Rows[i]["calculo"].ToString();

                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + valx + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + (totx * -1).ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{
                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}


                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }

                        //----------29/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION NACIONAL
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "DOLARES")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'N', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) / Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'N', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }


                        }
                        //---------- FIN 29/11/2017----------//

                        //----------13/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (prex > 0)
                        //    {
                        //        query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', 'A', 'N', 'LQ')";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();

                        //        query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();
                        //    }
                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 13/09/2017----------//


                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "INSERT INTO tb_det_prestamo (Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo) VALUES ('" + lblIdPrestamo.Text + "', '" + lbl_liquidacion.Text + "'," +
                        //        " '" + totx.ToString("0.00") + "', 'A', 'N')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }
                }

                cuerpo += "</table>";

                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);

                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }

                if (txtTL.Text == "PTC")
                {

                    /*SELECT A.hrp_recibo, A.hrp_num_unidades, B.prov_clave, A.prod_clave 
                        FROM tb_hist_recepcion A
                        INNER JOIN tb_mstr_recepcion_pt B ON B.rpt_recibo = A.hrp_recibo AND B.rpt_fecha = A.hrp_fecha AND B.rpt_tipo = A.hrp_situacion
                        WHERE hrp_fecha >= '01/06/2025'
                        AND A.hrp_fecha <= '30/06/2025' AND A.hrp_estatus <> 'C' AND A.hrp_numliq <> '' AND A.hrp_situacion = 'CM'
                        AND A.hrp_tipo_recepcion = 'PTC'  AND A.prod_clave = '16TOHBME10'*/

                    SqlDataAdapter adap = new SqlDataAdapter("SELECT A.hrp_recibo, A.hrp_num_unidades, B.prov_clave, A.prod_clave " +
                        "FROM tb_hist_recepcion A " +
                        "INNER JOIN tb_mstr_recepcion_pt B ON B.rpt_recibo = A.hrp_recibo AND B.rpt_tipo = A.hrp_situacion " +
                        "WHERE hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                        "AND A.hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.hrp_estatus <> 'C' AND A.hrp_numliq = '' AND A.hrp_situacion = 'CM' " +
                        "AND A.hrp_tipo_recepcion = 'PTC'  AND A.prod_clave = '" + lbl_cveprod.Text + "' AND B.prov_clave = '" + lbl_cveprov.Text + "'", thisConnection);
                    DataTable dtRecibs = new DataTable();
                    adap.Fill(dtRecibs);

                    foreach (DataRow row in dtRecibs.Rows)
                    {
                        string recib = row["hrp_recibo"].ToString();
                        string quant = row["hrp_num_unidades"].ToString();
                        query = "UPDATE tb_hist_recepcion SET hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' " +
                            "AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                            "AND hrp_estatus <> 'C' AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM' " +
                            "AND hrp_recibo = '" + recib + "'";
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();
                        cmnd3.Dispose();

                        query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();
                        cmnd3.Dispose();
                    }

                    dtRecibs = new DataTable();
                    adap = new SqlDataAdapter("SELECT A.hrp_recibo, A.hrp_num_unidades, B.prov_clave, A.prod_clave " +
                        "FROM tb_hist_recepcion A " +
                        "INNER JOIN tb_mstr_recepcion_pt B ON B.rpt_recibo = A.hrp_recibo AND B.rpt_tipo = A.hrp_situacion " +
                        "WHERE hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                        "AND A.hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.hrp_estatus <> 'C' AND A.hrp_numliq = '' AND A.hrp_situacion = 'CM' " +
                        "AND A.hrp_tipo_recepcion = 'PTC'  AND A.prod_clave = '" + lbl_cveprod.Text + "' AND B.prov_clave = '" + lbl_cveprov.Text + "'", thisConnection);
                    dtRecibs = new DataTable();
                    adap.Fill(dtRecibs);

                    if (dtRecibs.Rows.Count > 0)
                    {
                        foreach (DataRow row in dtRecibs.Rows)
                        {
                            string recib = row["hrp_recibo"].ToString();
                            string quant = row["hrp_num_unidades"].ToString();
                            query = "UPDATE tb_hist_recepcion SET hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' " +
                                "AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                                "AND hrp_estatus <> 'C' AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM' " +
                                "AND hrp_recibo = '" + recib + "'";
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = query;
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();

                            query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                                "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = query;
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();
                        }
                    }

                    //modificacion 06/11/2018
                    //asignacion de liquidacion a recibos de historico
                    //COMENTADO PARA REALIZAR NUEVO FORMATO DE ACTUALIZACION DE RECIBOS 08/07/2025
                    try
                    {
                        #region afectacion_de_recibos
                        //query = "SELECT A.rpt_recibo, B.rptd_cantidad FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                        //    "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                        //    "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                        //    "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = query;
                        //reader2 = cmnd2.ExecuteReader();
                        //if (reader2.HasRows)
                        //{
                        //    while (reader2.Read())
                        //    {


                        //        string recib = reader2["rpt_recibo"].ToString();
                        //        string quant = reader2["rptd_cantidad"].ToString();
                        //        query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                        //            " AND hrp_estatus <> 'C'  AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM'" +
                        //            " AND hrp_recibo = '" + recib + "' and (hrp_liquidado = 'F' OR hrp_liquidado = '')";
                        //        cmnd3 = thisConnection.CreateCommand();
                        //        cmnd3.CommandText = query;
                        //        cmnd3.ExecuteNonQuery();
                        //        cmnd3.Dispose();

                        //        query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                        //            "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                        //        cmnd3 = thisConnection.CreateCommand();
                        //        cmnd3.CommandText = query;
                        //        cmnd3.ExecuteNonQuery();
                        //        cmnd3.Dispose();

                        //    }
                        //}
                        //reader2.Close();
                        //reader1.Dispose();
                        //cmnd2.Dispose();

                        //DataTable dtOriginal = new DataTable();
                        //dtOriginal.Columns.Add("RPT_RECIBO", typeof(Int32));
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = "SELECT COUNT(A.rpt_recibo) AS CONTEO_RECIBOS FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                        //    "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                        //    "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                        //    "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                        //reader2 = cmnd2.ExecuteReader();
                        //DataRow RY;
                        //if (reader2.HasRows)
                        //{
                        //    reader2.Read();
                        //    RY = dtOriginal.NewRow();
                        //    RY["RPT_RECIBO"] = reader2["CONTEO_RECIBOS"].ToString();
                        //    dtOriginal.Rows.Add(RY);
                        //}
                        //reader2.Close();
                        //reader1.Dispose();
                        //cmnd2.Dispose();

                        //DataTable dtHistorico = new DataTable();
                        //dtHistorico.Columns.Add("HRP_RECIBO", typeof(Int32));

                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = "SELECT COUNT(hrp_recibo) AS CONTEO_HISTORICO FROM tb_hist_recepcion WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        //reader2 = cmnd2.ExecuteReader();
                        //DataRow RX;
                        //if (reader2.HasRows)
                        //{
                        //    reader2.Read();
                        //    RX = dtHistorico.NewRow();
                        //    RX["HRP_RECIBO"] = reader2["CONTEO_HISTORICO"].ToString();
                        //    dtHistorico.Rows.Add(RX);
                        //}
                        //reader2.Close();
                        //reader1.Dispose();
                        //cmnd2.Dispose();

                        //if (Convert.ToInt32(dtOriginal.Rows[0][0].ToString()) == Convert.ToInt32(dtHistorico.Rows[0][0].ToString()))
                        //{
                        //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + "TODOS LOS RECIBOS FUERON AFECTADOS " + lbl_liquidacion.Text);
                        //}
                        //else
                        //{
                        //    DataTable dtHistorico2 = new DataTable();
                        //    dtHistorico2.Columns.Add("HRP_RECIBO", typeof(Int32));
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = "SELECT hrp_recibo FROM tb_hist_recepcion WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        //    reader2 = cmnd2.ExecuteReader();
                        //    DataRow RXz;
                        //    if (reader2.HasRows)
                        //    {
                        //        reader2.Read();
                        //        RXz = dtHistorico2.NewRow();
                        //        RXz["HRP_RECIBO"] = reader2["hrp_recibo"].ToString();
                        //        dtHistorico2.Rows.Add(RXz);
                        //    }
                        //    reader2.Close();
                        //    reader1.Dispose();
                        //    cmnd2.Dispose();

                        //    string cadena_no_afectados = "";
                        //    query = "SELECT A.rpt_recibo, B.rptd_cantidad FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                        //    "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                        //    "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                        //    "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    reader2 = cmnd2.ExecuteReader();
                        //    if (reader2.HasRows)
                        //    {
                        //        while (reader2.Read())
                        //        {
                        //            string recib = reader2["rpt_recibo"].ToString();
                        //            string quant = reader2["rptd_cantidad"].ToString();
                        //            bool fnd = false;
                        //            foreach (DataRow rz in dtHistorico2.Select("HRP_RECIBO = '" + recib + "'"))
                        //            {
                        //                fnd = true;
                        //            }
                        //            if (fnd == false)
                        //            {
                        //                query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                        //                " AND hrp_estatus <> 'C'  AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM'" +
                        //                " AND hrp_recibo = '" + recib + "' and (hrp_liquidado = 'F' OR hrp_liquidado = '')";
                        //                cmnd3 = thisConnection.CreateCommand();
                        //                cmnd3.CommandText = query;
                        //                cmnd3.ExecuteNonQuery();
                        //                cmnd3.Dispose();

                        //                //cadena_no_afectados += cadena_no_afectados + " " + recib;

                        //                query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                        //                    "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                        //                cmnd3 = thisConnection.CreateCommand();
                        //                cmnd3.CommandText = query;
                        //                cmnd3.ExecuteNonQuery();
                        //                cmnd3.Dispose();
                        //            }
                        //        }
                        //    }
                        //    reader2.Close();
                        //    reader1.Dispose();
                        //    cmnd2.Dispose();

                        //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + "NO TODOS LOS RECIBOS FUERON AFECTADOS " + lbl_liquidacion.Text + " " + cadena_no_afectados);
                        //}
                        #endregion
                    }
                    catch (SqlException sqlex)
                    {
                        //MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //if (thisConnection.State == ConnectionState.Open)
                        //    thisConnection.Close();
                        //Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        //Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        ////this.Close();                                
                        //return;
                    }
                    //FIN COMENTADO PARA REALIZAR NUEVO FORMATO DE ACTUALIZACION DE RECIBOS 08/07/2025

                    //string qry = "";
                    //string var_recibo = "";
                    //try
                    //{
                    //    query = "SELECT hrp_recibo FROM tb_hist_recepcion WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "')" +
                    //        " AND hrp_estatus <> 'C'  AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion BETWEEN 'CM' AND 'CM')";//AND (lin_clave BETWEEN '" + txt_lincve.Text + "' AND '" + txt_lincve.Text + "')
                    //    cmnd2 = thisConnection.CreateCommand();
                    //    cmnd2.CommandText = query;
                    //    reader2 = cmnd2.ExecuteReader();
                    //    string var_cve_prov = "";
                    //    while (reader2.Read())
                    //    {
                    //        var_recibo = reader2.GetValue(0).ToString().Trim();

                    //        query = "SELECT prov_clave FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + var_recibo + "'";
                    //        cmnd3 = thisConnection.CreateCommand();
                    //        cmnd3.CommandText = query;
                    //        reader3 = cmnd3.ExecuteReader();
                    //        while (reader3.Read())
                    //        {
                    //            var_cve_prov = reader3.GetValue(0).ToString().Trim();
                    //        }
                    //        reader3.Close();
                    //        reader3.Dispose();
                    //        if (var_cve_prov == lbl_cveprov.Text)
                    //        {
                    //            query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "')" +
                    //                " AND hrp_estatus <> 'C'  AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion BETWEEN 'CM' AND 'CM')" +//AND (lin_clave BETWEEN '" + txt_lincve.Text + "' AND '" + txt_lincve.Text + "')
                    //                " AND hrp_recibo = '" + var_recibo + "' and hrp_liquidado = ''";
                    //            //cmnd3.CommandText = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_recibo = '" + var_recibo + "'";
                    //            cmnd3.CommandText = query;
                    //            cmnd3.ExecuteNonQuery();
                    //        }
                    //        cmnd3.Dispose();
                    //    }
                    //    reader2.Close();
                    //    reader2.Dispose();
                    //    cmnd2.Dispose();
                    //}
                    //catch (SqlException sqlex)
                    //{
                    //    MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    if (thisConnection.State == ConnectionState.Open)
                    //        thisConnection.Close();
                    //    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                    //    //this.Close();                                
                    //    return;
                    //}

                    ////foreach (DataRow row in tbrecibos.Rows)
                    ////{
                    ////    string var_recibo = row["hrp_recibo"].ToString();
                    ////    cmnd3 = thisConnection.CreateCommand();
                    ////    query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                    ////            " AND hrp_estatus <> 'C' AND lin_clave = '" + this.cvelin + "' AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion = 'CM')" +
                    ////            " AND hrp_recibo = '" + var_recibo + "'";
                    ////    //cmnd3.CommandText = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_recibo = '" + var_recibo + "'";
                    ////    cmnd3.CommandText = query;
                    ////    cmnd3.ExecuteNonQuery();
                    ////}

                }

                if (txtTL.Text == "PRO")
                {
                    string mprov = "";
                    string mprod = "";
                    string mliq = "";
                    string fliq = "";
                    mprov = lbl_cveprov.Text;
                    mprod = lbl_cveprod.Text;
                    mliq = lbl_liquidacion.Text;
                    fliq = lbl_fecha2.Text;
                    //DataRow dr;
                    //string ord_act = "";
                    //string ordp_ant = "";
                    foreach (DataRow rz in dtrecibos.Rows)
                    {
                        string qry = "";
                        try
                        {
                            string val = "";
                            if (txt_tipo.Text == "NACIONAL")
                                val = "N";
                            if (txt_tipo.Text == "EXPORTACION")
                                val = "E";
                            qry = "INSERT INTO tb_det_liq_planta(ordp_folio, prod_cve, prov_cve, liq_folio, liq_cant, liq_fecha, liq_noe, estatus) VALUES (" +
                                " '" + rz[0].ToString() + "', '" + mprod + "', '" + mprov + "', '" + mliq + "', '" + Convert.ToDecimal(rz[15].ToString()) + "', '" + Convert.ToDateTime(fliq).ToShortDateString() + "'," +
                                " '" + val + "', 'A')";
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = qry;
                            cmnd2.ExecuteNonQuery();
                            cmnd2.Dispose();

                            //qry = "UPDATE tb_hist_recepcion";
                            //cmnd2 = thisConnection.CreateCommand();
                            //cmnd2.CommandText = qry;
                            //cmnd2.ExecuteNonQuery();
                            //cmnd2.Dispose();
                        }
                        catch (SqlException sqlex)
                        {
                            MessageBox.Show("Error de sistema, no se termino de guardar la liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            if (thisConnection.State == ConnectionState.Open)
                                thisConnection.Close();
                            Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + qry, "SISEMP");
                            Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                            //this.Close();                                
                            return;
                        }
                    }
                }
                btnGuarda.Enabled = true;

                //27/07/2021
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                if (lblOrdenCompra.Text != "-")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8 FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    bool fnd = false;
                    string campo = "";
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            if (reader1["liq_numoc1"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc1";
                                break;
                            }
                            if (reader1["liq_numoc2"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc2";
                                break;
                            }
                            if (reader1["liq_numoc3"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc3";
                                break;
                            }
                            if (reader1["liq_numoc4"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc4";
                                break;
                            }
                            if (reader1["liq_numoc5"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc5";
                                break;
                            }
                            if (reader1["liq_numoc6"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc6";
                                break;
                            }
                            if (reader1["liq_numoc7"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc7";
                                break;
                            }
                            if (reader1["liq_numoc8"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc8";
                                break;
                            }
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (fnd == true)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET " + campo + " = '" + lblOrdenCompra.Text + "' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR RECIBOS DE HISTORICO
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numoc = '" + lblOrdenCompra.Text + "' WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "INSERT INTO tb_det_anticipada_pt (liq_folio, numero_oc, liq_cantidad, liq_fecha, liq_tipo) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + lblOrdenCompra.Text + "', '" + cantidad.ToString() + "', '" + DateTime.Now.ToShortDateString() + "', 'NACIONAL')";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR ORDEN DE COMPRA CAMPO liquidacion //actualizar si el campo es el liq_numoc1
                        if (campo == "liq_numoc1")
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '" + lbl_liquidacion.Text + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + cantidad.ToString() + "', unidad_oc = '" + lbl_liquidacion.Text + "' " +
                            "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //if (MessageBox.Show("¿Desea recalcular la orden de compra?", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        //{
                        //    if (MessageBox.Show("¿En verdad desea recalcular la orden de compra?\nYa no podrá surtir esta orden porque se cerrará.\nTendrá que cancelar la Orden de Compra y Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                        //    {
                        //        cmnd1 = thisConnection.CreateCommand();
                        //        cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND conse <> '" + lblConse.Text + "'";
                        //        decimal importe_det = 0;
                        //        reader1 = cmnd1.ExecuteReader();
                        //        if (reader1.HasRows)
                        //        {
                        //            reader1.Read();
                        //            importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                        //        }
                        //        reader1.Close();
                        //        reader1.Dispose();
                        //        cmnd1.Dispose();

                        //        cmnd1 = thisConnection.CreateCommand();
                        //        cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                        //            "cantidad_oc = cantidad_oc + '" + var_dec_unidades + "', " +
                        //            "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //            "importe_oc = importe_oc + '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                        //            "subtotal_oc = subtotal_oc + '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                        //            "total_oc = total_oc + '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                        //            "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //        cmnd1.ExecuteNonQuery();
                        //        cmnd1.Dispose();

                        //        cmnd1 = thisConnection.CreateCommand();
                        //        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                        //            "cantidad_oc = '" + var_dec_unidades + "', " +
                        //            "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //            "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                        //            "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        //        cmnd1.ExecuteNonQuery();
                        //        cmnd1.Dispose();
                        //    }   
                        //}

                        //if (chkRecalculo.Checked == true)
                        //{
                        //RECALCULO DE ORDEN DE COMPRA

                        //}
                        MessageBox.Show("Si es la última recepcion, favor de recalcular consultando la Orden de Compra de PT", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }



                }
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                //27/07/2021

                thisConnection.Close();

                afecta_notas_credito_nacional(lbl_liquidacion.Text, lbl_fecha1.Text, lbl_fecha2.Text, lbl_cveprod.Text, lbl_cveprov.Text);
                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //printDocument1.Print();
                }


                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                //Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");

                MessageBox.Show("Datos Guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (chkRecalculo.Checked == true)
                {
                    MessageBox.Show("Se recalculo la orden de compra, favor de consultarla e imprimirla nuevamente", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void guardanuevoexp()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;



            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");
                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                //rr["moni"] = dtgConceptos.Rows[i].Cells["mon"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }

            string cvep = "";
            string var_chr_prod_clave = "";
            decimal canti = 0;

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;
            canti = 0;

            if (var_chr_prod_clave == "05005LETOR" || var_chr_prod_clave == "05005LETOT")
            {
                if (tcon.Rows[0][0].ToString() == "1")
                {
                    canti = Math.Round(Convert.ToDecimal(lbl_libras.Text) * Convert.ToDecimal(tcon.Rows[0]["unidades"].ToString()), 2);
                }
            }

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT COUNT(liq_folio) AS conteo FROM tb_mstr_liquidacion WHERE liq_fecha1 = '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                "AND liq_fecha2 = '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND liq_provcve = '" + lbl_cveprov.Text + "' " +
                "AND liq_lincve = '" + txt_lincve.Text + "' AND liq_prodcve = '" + lbl_cveprod.Text + "' AND uni_exp = '" + tcon.Rows[0]["unidades"].ToString() + "' and status = 'A'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                Int32 val = Convert.ToInt32(reader1["conteo"].ToString());
                //if (reader1["conteo"].ToString() == "" || reader1["conteo"].ToString() == null || reader1["conteo"].ToString() == "0")
                //{
                //    //return;
                //}
                //else 
                //{
                //    val = Convert.ToInt32(reader1["conteo"].ToString());
                //}
                if (val > 0)
                {
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                    return;
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();

                query = "INSERT INTO tb_mstr_liquidacion" +
                        "(liq_fecha1, liq_fecha2, liq_provcve, liq_lincve," +//1
                        " liq_linnom, liq_provnom, liq_prodcve, liq_prodnom, liq_cajas_pal," +//2
                        " liq_flejes_pal, liq_unidades, liq_mermas, uni_nac, uninac_oc," +//3
                        " uni_mern, uni_exp, uniexp_oc, uni_mere, liq_pre_uni, " +//4
                        " liq_porcen1, liq_porcen2, liq_imp_tot, liq_imp_por, liq_imp_liq," +//5
                        " liq_costo1, liq_costo2, liq_afecto, liq_precambio, liq_preunie," +//6
                        " liq_porcen_1e, liq_porcen_2e, liq_imp_tote, liq_imp_pore, liq_imp_liqe," +//7
                        " liq_costo1e, liq_costo2e, status, liq_ocompra, liq_numoc," +//8
                        " liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8," +//9
                        " liq_cantiocn, liq_cantioce, liq_libras, liq_por_des, liq_imp_pordes," +//10
                        " liq_exp_pordes, liq_exp_imppordes, tipo, liq_fecha, liq_exp, conse)" +//11
                        " VALUES" +
                        " ('" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "', '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "', '" + lbl_cveprov.Text + "', '" + txt_lincve.Text + "'," +
                        " '" + txt_linnom.Text + "', '" + lbl_proveedor.Text + "', '" + lbl_cveprod.Text + "', '" + lbl_producto.Text.Replace("'", " ").ToString() + "', '" + lbl_cajas.Text + "'," +
                        " '" + lbl_flejes.Text + "', '" + cantidad + "', 0, 0, 0," +
                        " 0, '" + var_dec_unidades + "', 0, 0, 0," +//4
                        " 0, 0, 0, 0, 0," +//5
                        " 0, 0, '0', '" + Convert.ToDecimal(txt_tipocambio.Text).ToString("0.0000") + "', '" + var_dec_precio + "'," +//6
                        " '" + var_dec_prod_comision + "', '" + var_dec_prod_comision2 + "', '" + Convert.ToDecimal(txt_total.Text).ToString("0.0000") + "', '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.0000") + "', '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.0000") + "'," +//7
                        " '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.0000") + "', '" + ((txt_nuevocosto.Text == "") ? txt_nuevocosto.Text : "0") + "', 'A', '', ''," +//8
                        " '', '', '', '', '', '', '', ''," +
                        " 0, 0, '" + canti + "', 0, 0," +//10
                        " '" + txt_porce_desc.Text + "', '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.0000") + "', " +
                        "'" + tl + "', '" + DateTime.Now.ToShortDateString() + "', '" + lblTeorico.Text + "', '" + ((lblConse.Text == "") ? "0" : lblConse.Text) + "') SELECT SCOPE_IDENTITY()";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                if (lbl_liquidacion.Text != ultimo_folio)
                {
                    MessageBox.Show("El folio de la liquidacion ha cambiado por movimientos en la red, el numero de folio asignado es: " + ultimo_folio + "\nSe imprimirá la nueva liquidacion enseguida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    lbl_liquidacion.Text = ultimo_folio;
                    //printDocument1.Print();
                    thisConnection.Close();
                    //printDocument1.Print();
                    thisConnection.Open();
                }
                else
                {
                    thisConnection.Close();
                    //printDocument1.Print();
                    thisConnection.Open();
                }

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Inserción de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", lbl_liquidacion.Text, "INSERCION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();
                    calc = tcon.Rows[i]["calculo"].ToString();

                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + lblIdPrestamo.Text + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{


                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                                " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}

                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }

                        //----------29/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION NACIONAL
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "PESOS")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'E', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) * Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'E', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }
                        }
                        //---------- FIN 29/11/2017----------//

                        //----------13/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (prex > 0)
                        //    {
                        //        query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', 'A', 'E', 'LQ')";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();

                        //        query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();
                        //    }

                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 13/09/2017----------//

                        //query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E')";
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = query;
                        //cmnd2.ExecuteNonQuery();
                        //cmnd2.Dispose();

                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "INSERT INTO tb_det_prestamo (Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo) VALUES ('" + lblIdPrestamo.Text + "', '" + lbl_liquidacion.Text + "'," +
                        //        " '" + totx.ToString("0.00") + "', 'A', 'E')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }
                }

                cuerpo += "</table>";
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);

                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }

                if (txtTL.Text == "PTC")
                {

                    //asignacion de liquidacion a recibos de historico

                    SqlDataAdapter adap = new SqlDataAdapter("SELECT A.hrp_recibo, A.hrp_num_unidades, B.prov_clave, A.prod_clave " +
                        "FROM tb_hist_recepcion A " +
                        "INNER JOIN tb_mstr_recepcion_pt B ON B.rpt_recibo = A.hrp_recibo AND B.rpt_tipo = A.hrp_situacion " +
                        "WHERE hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                        "AND A.hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.hrp_estatus <> 'C' AND A.hrp_numliq = '' AND A.hrp_situacion = 'CM' " +
                        "AND A.hrp_tipo_recepcion = 'PTC'  AND A.prod_clave = '" + lbl_cveprod.Text + "' AND B.prov_clave = '" + lbl_cveprov.Text + "'", thisConnection);
                    DataTable dtRecibs = new DataTable();
                    adap.Fill(dtRecibs);

                    foreach (DataRow row in dtRecibs.Rows)
                    {
                        string recib = row["hrp_recibo"].ToString();
                        string quant = row["hrp_num_unidades"].ToString();
                        query = "UPDATE tb_hist_recepcion SET hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' " +
                            "AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                            "AND hrp_estatus <> 'C' AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM' " +
                            "AND hrp_recibo = '" + recib + "'";
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();
                        cmnd3.Dispose();

                        query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                        cmnd3 = thisConnection.CreateCommand();
                        cmnd3.CommandText = query;
                        cmnd3.ExecuteNonQuery();
                        cmnd3.Dispose();
                    }

                    dtRecibs = new DataTable();
                    adap = new SqlDataAdapter("SELECT A.hrp_recibo, A.hrp_num_unidades, B.prov_clave, A.prod_clave " +
                        "FROM tb_hist_recepcion A " +
                        "INNER JOIN tb_mstr_recepcion_pt B ON B.rpt_recibo = A.hrp_recibo AND B.rpt_tipo = A.hrp_situacion " +
                        "WHERE hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' " +
                        "AND A.hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.hrp_estatus <> 'C' AND A.hrp_numliq = '' AND A.hrp_situacion = 'CM' " +
                        "AND A.hrp_tipo_recepcion = 'PTC'  AND A.prod_clave = '" + lbl_cveprod.Text + "' AND B.prov_clave = '" + lbl_cveprov.Text + "'", thisConnection);
                    dtRecibs = new DataTable();
                    adap.Fill(dtRecibs);

                    if (dtRecibs.Rows.Count > 0)
                    {
                        foreach (DataRow row in dtRecibs.Rows)
                        {
                            string recib = row["hrp_recibo"].ToString();
                            string quant = row["hrp_num_unidades"].ToString();
                            query = "UPDATE tb_hist_recepcion SET hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' " +
                                "AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                                "AND hrp_estatus <> 'C' AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM' " +
                                "AND hrp_recibo = '" + recib + "'";
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = query;
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();

                            query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                                "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                            cmnd3 = thisConnection.CreateCommand();
                            cmnd3.CommandText = query;
                            cmnd3.ExecuteNonQuery();
                            cmnd3.Dispose();
                        }
                    }

                    //COMENTADO PARA REALIZAR NUEVO FORMATO DE ACTUALIZACION DE RECIBOS 08/07/2025
                    //modificacion 06/11/2018
                    #region afectacion_de_recibos
                    //try
                    //{
                    //    query = "SELECT A.rpt_recibo, B.rptd_cantidad FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                    //        "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                    //        "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                    //        "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                    //    cmnd2 = thisConnection.CreateCommand();
                    //    cmnd2.CommandText = query;
                    //    reader2 = cmnd2.ExecuteReader();
                    //    if (reader2.HasRows)
                    //    {
                    //        while (reader2.Read())
                    //        {
                    //            string recib = reader2["rpt_recibo"].ToString();
                    //            string quant = reader2["rptd_cantidad"].ToString();
                    //            query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                    //                " AND hrp_estatus <> 'C'  AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM'" +
                    //                " AND hrp_recibo = '" + recib + "' and (hrp_liquidado = 'F' OR hrp_liquidado = '')";
                    //            cmnd3 = thisConnection.CreateCommand();
                    //            cmnd3.CommandText = query;
                    //            cmnd3.ExecuteNonQuery();
                    //            cmnd3.Dispose();

                    //            query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                    //                        "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                    //            cmnd3 = thisConnection.CreateCommand();
                    //            cmnd3.CommandText = query;
                    //            cmnd3.ExecuteNonQuery();
                    //            cmnd3.Dispose();
                    //        }
                    //    }
                    //    reader2.Close();
                    //    reader1.Dispose();
                    //    cmnd2.Dispose();

                    //    DataTable dtOriginal = new DataTable();
                    //    dtOriginal.Columns.Add("RPT_RECIBO", typeof(Int32));
                    //    cmnd2 = thisConnection.CreateCommand();
                    //    cmnd2.CommandText = "SELECT COUNT(A.rpt_recibo) AS CONTEO_RECIBOS FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                    //        "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                    //        "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                    //        "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                    //    reader2 = cmnd2.ExecuteReader();
                    //    DataRow RY;
                    //    if (reader2.HasRows)
                    //    {
                    //        reader2.Read();
                    //        RY = dtOriginal.NewRow();
                    //        RY["RPT_RECIBO"] = reader2["CONTEO_RECIBOS"].ToString();
                    //        dtOriginal.Rows.Add(RY);
                    //    }
                    //    reader2.Close();
                    //    reader1.Dispose();
                    //    cmnd2.Dispose();

                    //    DataTable dtHistorico = new DataTable();
                    //    dtHistorico.Columns.Add("HRP_RECIBO", typeof(Int32));
                    //    cmnd2 = thisConnection.CreateCommand();
                    //    cmnd2.CommandText = "SELECT COUNT(hrp_recibo) AS CONTEO_HISTORICO FROM tb_hist_recepcion WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                    //    reader2 = cmnd2.ExecuteReader();
                    //    DataRow RX;
                    //    if (reader2.HasRows)
                    //    {
                    //        reader2.Read();
                    //        RX = dtHistorico.NewRow();
                    //        RX["HRP_RECIBO"] = reader2["CONTEO_HISTORICO"].ToString();
                    //        dtHistorico.Rows.Add(RX);
                    //    }
                    //    reader2.Close();
                    //    reader1.Dispose();
                    //    cmnd2.Dispose();

                    //    if (dtOriginal.Rows.Count == dtHistorico.Rows.Count)
                    //    {
                    //        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + "TODOS LOS RECIBOS FUERON AFECTADOS " + lbl_liquidacion.Text);
                    //    }
                    //    else
                    //    {
                    //        string cadena_no_afectados = "";
                    //        query = "SELECT A.rpt_recibo, B.rptd_cantidad FROM tb_mstr_recepcion_pt A JOIN tb_det_recepcion_pt B ON A.rpt_recibo = B.rpt_recibo " +
                    //        "WHERE A.rpt_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND " +
                    //        "A.rpt_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' AND A.rpt_tipo = 'CM' AND A.prov_clave = '" + lbl_cveprov.Text + "' AND " +
                    //        "A.rpt_estatus = '' AND B.prod_clave = '" + lbl_cveprod.Text + "'";
                    //        cmnd2 = thisConnection.CreateCommand();
                    //        cmnd2.CommandText = query;
                    //        reader2 = cmnd2.ExecuteReader();
                    //        if (reader2.HasRows)
                    //        {
                    //            while (reader2.Read())
                    //            {
                    //                string recib = reader2["rpt_recibo"].ToString();
                    //                string quant = reader2["rptd_cantidad"].ToString();
                    //                bool fnd = false;
                    //                foreach (DataRow rz in dtHistorico.Select("HRP_RECIBO = '" + recib + "'"))
                    //                {
                    //                    fnd = true;
                    //                }
                    //                if (fnd == false)
                    //                {
                    //                    query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND hrp_fecha >= '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND hrp_fecha <= '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "' " +
                    //                    " AND hrp_estatus <> 'C'  AND prod_clave = '" + lbl_cveprod.Text + "' AND hrp_situacion = 'CM'" +
                    //                    " AND hrp_recibo = '" + recib + "' and (hrp_liquidado = 'F' OR hrp_liquidado = '')";
                    //                    cmnd3 = thisConnection.CreateCommand();
                    //                    cmnd3.CommandText = query;
                    //                    cmnd3.ExecuteNonQuery();
                    //                    cmnd3.Dispose();

                    //                    //cadena_no_afectados += cadena_no_afectados + " " + recib;

                    //                    query = "INSERT INTO tb_det_liquidacion_rec(liquidacion, recibo, cantidad, status, producto) " +
                    //                        "VALUES('" + lbl_liquidacion.Text + "', '" + recib + "', '" + quant + "', 'A', '" + lbl_cveprod.Text + "')";
                    //                    cmnd3 = thisConnection.CreateCommand();
                    //                    cmnd3.CommandText = query;
                    //                    cmnd3.ExecuteNonQuery();
                    //                    cmnd3.Dispose();
                    //                }
                    //            }
                    //        }
                    //        reader2.Close();
                    //        reader1.Dispose();
                    //        cmnd2.Dispose();

                    //        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + "NO TODOS LOS RECIBOS FUERON AFECTADOS " + lbl_liquidacion.Text + " " + cadena_no_afectados);
                    //    }

                    //}
                    //catch (SqlException sqlex)
                    //{
                    //    MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    if (thisConnection.State == ConnectionState.Open)
                    //        thisConnection.Close();
                    //    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                    //    //this.Close();                                
                    //    return;
                    //}

                    #endregion
                    //FIN COMENTADO PARA REALIZAR NUEVO FORMATO DE ACTUALIZACION DE RECIBOS 08/07/2025

                    //string qry = "";
                    //string var_recibo = "";
                    //try
                    //{
                    //    query = "SELECT hrp_recibo FROM tb_hist_recepcion WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "')" +
                    //        " AND hrp_estatus <> 'C' AND (lin_clave BETWEEN '" + txt_lincve.Text + "' AND '" + txt_lincve.Text + "') AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion BETWEEN 'CM' AND 'CM')";
                    //    cmnd2 = thisConnection.CreateCommand();
                    //    cmnd2.CommandText = query;
                    //    reader2 = cmnd2.ExecuteReader();
                    //    string var_cve_prov = "";
                    //    while (reader2.Read())
                    //    {
                    //        var_recibo = reader2.GetValue(0).ToString().Trim();

                    //        query = "SELECT prov_clave FROM tb_mstr_recepcion_pt WHERE rpt_recibo = '" + var_recibo + "'";
                    //        cmnd3 = thisConnection.CreateCommand();
                    //        cmnd3.CommandText = query;
                    //        reader3 = cmnd3.ExecuteReader();
                    //        while (reader3.Read())
                    //        {
                    //            var_cve_prov = reader3.GetValue(0).ToString().Trim();
                    //        }
                    //        reader3.Close();
                    //        reader3.Dispose();
                    //        if (var_cve_prov == lbl_cveprov.Text)
                    //        {
                    //            query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(lbl_fecha1.Text).ToShortDateString() + "' AND '" + Convert.ToDateTime(lbl_fecha2.Text).ToShortDateString() + "')" +
                    //                " AND hrp_estatus <> 'C' AND (lin_clave BETWEEN '" + txt_lincve.Text + "' AND '" + txt_lincve.Text + "') AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion BETWEEN 'CM' AND 'CM')" +
                    //                " AND hrp_recibo = '" + var_recibo + "'  and hrp_liquidado = 'F'";
                    //            //cmnd3.CommandText = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_recibo = '" + var_recibo + "'";
                    //            cmnd3.CommandText = query;
                    //            cmnd3.ExecuteNonQuery();
                    //        }
                    //        cmnd3.Dispose();
                    //    }
                    //    reader2.Close();
                    //    reader2.Dispose();
                    //    cmnd2.Dispose();
                    //}
                    //catch (SqlException sqlex)
                    //{
                    //    MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    if (thisConnection.State == ConnectionState.Open)
                    //        thisConnection.Close();
                    //    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    //    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                    //    //this.Close();                                
                    //    return;
                    //}
                    ////foreach (DataRow row in tbrecibos.Rows)
                    ////{
                    ////    //string var_recibo = row["hrp_recibo"].ToString();
                    ////    //cmnd3 = thisConnection.CreateCommand();
                    ////    //query = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_tipo_recepcion = 'PTC' AND (hrp_fecha BETWEEN '" + Convert.ToDateTime(f1).ToShortDateString() + "' AND '" + Convert.ToDateTime(f2).ToShortDateString() + "')" +
                    ////    //        " AND hrp_estatus <> 'C' AND lin_clave = '" + this.cvelin + "' AND (prod_clave BETWEEN '" + cvep + "' AND '" + cvep + "') AND (hrp_situacion = 'CM')" +
                    ////    //        " AND hrp_recibo = '" + var_recibo + "'";
                    ////    ////cmnd3.CommandText = "UPDATE tb_hist_recepcion SET hrp_liquidado = 'T', hrp_numliq = '" + lbl_liquidacion.Text + "' WHERE hrp_recibo = '" + var_recibo + "'";
                    ////    //cmnd3.CommandText = query;
                    ////    //cmnd3.ExecuteNonQuery();
                    ////}

                }

                if (txtTL.Text == "PRO")
                {
                    string mprov = "";
                    string mprod = "";
                    string mliq = "";
                    string fliq = "";
                    mprov = lbl_cveprov.Text;
                    mprod = lbl_cveprod.Text;
                    mliq = lbl_liquidacion.Text;
                    fliq = lbl_fecha2.Text;
                    DataRow dr;
                    string ord_act = "";
                    string ordp_ant = "";

                    foreach (DataRow rz in dtrecibos.Rows)
                    {
                        string qry = "";
                        try
                        {
                            string val = "";
                            if (txt_tipo.Text == "NACIONAL")
                                val = "N";
                            if (txt_tipo.Text == "EXPORTACION")
                                val = "E";
                            qry = "INSERT INTO tb_det_liq_planta(ordp_folio, prod_cve, prov_cve, liq_folio, liq_cant, liq_fecha, liq_noe, estatus) VALUES (" +
                                " '" + rz[0].ToString() + "', '" + mprod + "', '" + mprov + "', '" + mliq + "', '" + Convert.ToDecimal(rz[15].ToString()) + "', '" + Convert.ToDateTime(fliq).ToShortDateString() + "'," +
                                " '" + val + "', 'A')";
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = qry;
                            cmnd2.ExecuteNonQuery();
                            cmnd2.Dispose();

                            //qry = "UPDATE tb_hist_recepcion";
                            //cmnd2 = thisConnection.CreateCommand();
                            //cmnd2.CommandText = qry;
                            //cmnd2.ExecuteNonQuery();
                            //cmnd2.Dispose();
                        }
                        catch (SqlException sqlex)
                        {
                            MessageBox.Show("Error de sistema, no se termino de guardar la liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            if (thisConnection.State == ConnectionState.Open)
                                thisConnection.Close();
                            Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + qry, "SISEMP");
                            Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                            //this.Close();                                
                            return;
                        }
                    }
                    //foreach (DataRow rz in reciboz.Rows)
                    //{
                    //    string query = "";
                    //    try
                    //    {
                    //        query = "INSERT INTO tb_det_liq_planta(ordp_folio, prod_cve, prov_cve, liq_folio, liq_cant, liq_fecha, liq_noe, estatus) VALUES (" +
                    //            " '" + rz[0].ToString() + "', '" + mprod + "', '" + mprov + "', '" + mliq + "', '" + Convert.ToDecimal(rz[15].ToString()) + "', '" + Convert.ToDateTime(fliq).ToShortDateString() + "'," +
                    //            " '" + txt_tipo.Text + "', 'A')";
                    //        cmnd2 = thisConnection.CreateCommand();
                    //        cmnd2.CommandText = query;
                    //        cmnd2.ExecuteNonQuery();
                    //        cmnd2.Dispose();
                    //    }
                    //    catch (SqlException sqlex)
                    //    {
                    //        MessageBox.Show("Error de sistema, no se termino de guardar la liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //        if (thisConnection.State == ConnectionState.Open)
                    //            thisConnection.Close();
                    //        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                    //        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                    //        //this.Close();                                
                    //        return;
                    //    }
                    //}
                }
                btnGuarda.Enabled = true;

                //28/07/2021
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                if (lblOrdenCompra.Text != "-")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8 FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    bool fnd = false;
                    string campo = "";
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            if (reader1["liq_numoc1"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc1";
                                break;
                            }
                            if (reader1["liq_numoc2"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc2";
                                break;
                            }
                            if (reader1["liq_numoc3"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc3";
                                break;
                            }
                            if (reader1["liq_numoc4"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc4";
                                break;
                            }
                            if (reader1["liq_numoc5"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc5";
                                break;
                            }
                            if (reader1["liq_numoc6"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc6";
                                break;
                            }
                            if (reader1["liq_numoc7"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc7";
                                break;
                            }
                            if (reader1["liq_numoc8"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc8";
                                break;
                            }
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (fnd == true)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET " + campo + " = '" + lblOrdenCompra.Text + "' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR RECIBOS DE HISTORICO
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numoc = '" + lblOrdenCompra.Text + "' WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "INSERT INTO tb_det_anticipada_pt (liq_folio, numero_oc, liq_cantidad, liq_fecha, liq_tipo) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + lblOrdenCompra.Text + "', '" + cantidad.ToString() + "', '" + DateTime.Now.ToShortDateString() + "', 'EXPORTACION')";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR ORDEN DE COMPRA CAMPO liquidacion
                        if (campo == "liq_numoc1")
                        {
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '" + lbl_liquidacion.Text + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                            cmnd1.ExecuteNonQuery();
                            cmnd1.Dispose();
                        }


                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + cantidad.ToString() + "', unidad_oc = '" + lbl_liquidacion.Text + "' " +
                            "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        ////if (chkRecalculo.Checked == true)
                        ////{
                        //    //RECALCULO DE ORDEN DE COMPRA
                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND conse <> '" + lblConse.Text + "'";
                        //    decimal importe_det = 0;
                        //    reader1 = cmnd1.ExecuteReader();
                        //    if (reader1.HasRows)
                        //    {
                        //        reader1.Read();
                        //        importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                        //    }
                        //    reader1.Close();
                        //    reader1.Dispose();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                        //        "subtotal_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                        //        "total_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();
                        ////}

                        MessageBox.Show("Si es la última recepcion, favor de recalcular consultando la Orden de Compra de PT", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                //28/07/2021

                thisConnection.Close();

                //modificar notas de credito
                afecta_notas_credito_exportacion(lbl_liquidacion.Text, lbl_fecha1.Text, lbl_fecha2.Text, lbl_cveprod.Text, lbl_cveprov.Text);

                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //printDocument1.Print();
                }

                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                //Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");

                MessageBox.Show("Datos Guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void guardafaltantenal()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;
            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");

                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                //rr["moni"] = dtgConceptos.Rows[i].Cells["mon"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }

            string cvep = "";
            string var_chr_prod_clave = "";

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();
            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();

                query = "UPDATE tb_mstr_liquidacion SET uni_nac = '" + var_dec_unidades + "', liq_pre_uni = '" + var_dec_precio + "', liq_porcen1 = '" + var_dec_prod_comision + "', liq_porcen2 = '" + var_dec_prod_comision2 + "', " +
                    "liq_imp_tot = '" + Convert.ToDecimal(txt_total.Text).ToString("0.000") + "', liq_imp_por = '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.000") + "', liq_imp_liq = '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.000") + "', " +
                    "liq_costo1 = '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.000") + "', liq_costo2 = '" + ((txt_nuevocosto.Text == "") ? Convert.ToDecimal(txt_nuevocosto.Text).ToString("0.00") : "0") + "', " +
                    "liq_por_des = '" + Convert.ToDecimal(txt_porce_desc.Text).ToString("0.000") + "', liq_imp_pordes = '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.000") + "', status = 'A', liq_nac = '" + lblTeorico.Text + "' " +
                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                //ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                //if (lbl_liquidacion.Text != ultimo_folio)
                //{
                //    MessageBox.Show("El folio de la liquidacion ha cambiado por movimientos en la red, el numero de folio asignado es: " + ultimo_folio + "\nSe imprimirá la nueva liquidacion enseguida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    lbl_liquidacion.Text = ultimo_folio;
                //    //printDocument1.Print();
                //}
                //else
                //{
                ////printDocument1.Print();
                //}

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Modificación de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", lbl_liquidacion.Text, "MODIFICACION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text+ "' AND tipo_con = 'N'";
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();
                    calc = tcon.Rows[i]["calculo"].ToString();

                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + lblIdPrestamo.Text + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{
                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}

                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }

                        //----------30/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION NACIONAL
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "DOLARES")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'N', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) / Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'N', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }
                        }
                        //---------- FIN 30/11/2017----------//

                        //----------13/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (prex > 0)
                        //    {
                        //        query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + totx.ToString("0.0000") + "', 'A', 'N', 'LQ')";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();

                        //        query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();
                        //    }

                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 13/09/2017----------//

                        //query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N')";
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = query;
                        //cmnd2.ExecuteNonQuery();
                        //cmnd2.Dispose();

                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "INSERT INTO tb_det_prestamo (Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo) VALUES ('" + lblIdPrestamo.Text + "', '" + lbl_liquidacion.Text + "'," +
                        //        " '" + totx.ToString("0.00") + "', 'A', 'N')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }
                }

                cuerpo += "</table>";
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);

                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }

                btnGuarda.Enabled = true;

                //28/07/2021
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                if (lblOrdenCompra.Text != "-")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8 FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    bool fnd = false;
                    string campo = "";
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            if (reader1["liq_numoc1"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc1";
                                break;
                            }
                            if (reader1["liq_numoc2"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc2";
                                break;
                            }
                            if (reader1["liq_numoc3"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc3";
                                break;
                            }
                            if (reader1["liq_numoc4"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc4";
                                break;
                            }
                            if (reader1["liq_numoc5"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc5";
                                break;
                            }
                            if (reader1["liq_numoc6"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc6";
                                break;
                            }
                            if (reader1["liq_numoc7"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc7";
                                break;
                            }
                            if (reader1["liq_numoc8"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc8";
                                break;
                            }
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (fnd == true)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET " + campo + " = '" + lblOrdenCompra.Text + "' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR RECIBOS DE HISTORICO
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numoc = '" + lblOrdenCompra.Text + "' WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "INSERT INTO tb_det_anticipada_pt (liq_folio, numero_oc, liq_cantidad, liq_fecha, liq_tipo) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + lblOrdenCompra.Text + "', '" + cantidad.ToString() + "', '" + DateTime.Now.ToShortDateString() + "', 'NACIONAL')";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR ORDEN DE COMPRA CAMPO liquidacion
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '" + lbl_liquidacion.Text + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = '" + cantidad.ToString() + "', unidad_oc = '" + lbl_liquidacion.Text + "' " +
                            "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        ////if (chkRecalculo.Checked == true)
                        ////{
                        //    //RECALCULO DE ORDEN DE COMPRA
                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND conse <> '" + lblConse.Text + "'";
                        //    decimal importe_det = 0;
                        //    reader1 = cmnd1.ExecuteReader();
                        //    if (reader1.HasRows)
                        //    {
                        //        reader1.Read();
                        //        importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                        //    }
                        //    reader1.Close();
                        //    reader1.Dispose();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                        //        "subtotal_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                        //        "total_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();
                        ////}
                    }
                }
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                //28/07/2021

                thisConnection.Close();

                afecta_notas_credito_nacional(lbl_liquidacion.Text, lbl_fecha1.Text, lbl_fecha2.Text, lbl_cveprod.Text, lbl_cveprov.Text);

                MessageBox.Show("Datos Guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //printDocument1.Print();
                }

                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                ////Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void guardafaltanteexp()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;
            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");

                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                //rr["moni"] = dtgConceptos.Rows[i].Cells["mon"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }

            string cvep = "";
            string var_chr_prod_clave = "";
            decimal canti = 0;

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;
            canti = 0;

            if (var_chr_prod_clave == "05005LETOR" || var_chr_prod_clave == "05005LETOT")
            {
                if (tcon.Rows[0][0].ToString() == "1")
                {
                    canti = Math.Round(Convert.ToDecimal(lbl_libras.Text) * Convert.ToDecimal(tcon.Rows[0]["unidades"].ToString()), 2);
                }
            }

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();
            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();

                query = "UPDATE tb_mstr_liquidacion SET uni_exp = '" + var_dec_unidades + "', liq_preunie = '" + var_dec_precio + "', liq_porcen_1e = '" + var_dec_prod_comision + "', liq_porcen_2e = '" + var_dec_prod_comision2 + "', " +
                    "liq_imp_tote = '" + Convert.ToDecimal(txt_total.Text).ToString("0.000") + "', liq_imp_pore = '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.000") + "', liq_imp_liqe = '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.000") + "', " +
                    "liq_costo1e = '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.000") + "', liq_costo2e = '" + ((txt_nuevocosto.Text == "") ? Convert.ToDecimal(txt_nuevocosto.Text).ToString("0.00") : "0") + "', " +
                    "liq_exp_pordes = '" + Convert.ToDecimal(txt_porce_desc.Text).ToString("0.000") + "', liq_exp_imppordes = '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.000") + "', status = 'A', liq_precambio = '" + Convert.ToDecimal(txt_tipocambio.Text) + "', liq_exp = '" + lblTeorico.Text + "' " +
                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                //ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                //if (lbl_liquidacion.Text != ultimo_folio)
                //{
                //    MessageBox.Show("El folio de la liquidacion ha cambiado por movimientos en la red, el numero de folio asignado es: " + ultimo_folio + "\nSe imprimirá la nueva liquidacion enseguida", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    //lbl_liquidacion.Text = ultimo_folio;
                //    //printDocument1.Print();
                //}
                //else
                //{
                ////printDocument1.Print();
                //}

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Inserción de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", lbl_liquidacion.Text, "INSERCION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //cmnd1 = thisConnection.CreateCommand();
                //cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND tipo_con = 'E'";
                //cmnd1.ExecuteNonQuery();
                //cmnd1.Dispose();

                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();

                    calc = tcon.Rows[i]["calculo"].ToString();
                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + lblIdPrestamo.Text + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{
                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}
                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }

                        //----------30/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION NACIONAL
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "PESOS")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'E', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) * Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'E', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }
                        }
                        //---------- FIN 30/11/2017----------//

                        //----------13/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (prex > 0)
                        //    {
                        //        query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + totx.ToString("0.0000") + "', 'A', 'E', 'LQ')";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();

                        //        query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //        cmnd2 = thisConnection.CreateCommand();
                        //        cmnd2.CommandText = query;
                        //        cmnd2.ExecuteNonQuery();
                        //        cmnd2.Dispose();
                        //    }

                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 13/09/2017----------//
                        //query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E')";
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = query;
                        //cmnd2.ExecuteNonQuery();
                        //cmnd2.Dispose();

                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "INSERT INTO tb_det_prestamo (Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo) VALUES ('" + lblIdPrestamo.Text + "', '" + lbl_liquidacion.Text + "'," +
                        //        " '" + totx.ToString("0.00") + "', 'A', 'E')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }
                }

                cuerpo += "</table>";
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);

                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }

                btnGuarda.Enabled = true;

                //28/07/2021
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                if (lblOrdenCompra.Text != "-")
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8 FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    bool fnd = false;
                    string campo = "";
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            if (reader1["liq_numoc1"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc1";
                                break;
                            }
                            if (reader1["liq_numoc2"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc2";
                                break;
                            }
                            if (reader1["liq_numoc3"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc3";
                                break;
                            }
                            if (reader1["liq_numoc4"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc4";
                                break;
                            }
                            if (reader1["liq_numoc5"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc5";
                                break;
                            }
                            if (reader1["liq_numoc6"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc6";
                                break;
                            }
                            if (reader1["liq_numoc7"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc7";
                                break;
                            }
                            if (reader1["liq_numoc8"].ToString().Trim() == "")
                            {
                                fnd = true;
                                campo = "liq_numoc8";
                                break;
                            }
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (fnd == true)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET " + campo + " = '" + lblOrdenCompra.Text + "' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR RECIBOS DE HISTORICO
                        //cmnd1 = thisConnection.CreateCommand();
                        //cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numoc = '" + lblOrdenCompra.Text + "' WHERE hrp_numliq = '" + lbl_liquidacion.Text + "'";
                        //cmnd1.ExecuteNonQuery();
                        //cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "INSERT INTO tb_det_anticipada_pt (liq_folio, numero_oc, liq_cantidad, liq_fecha, liq_tipo) " +
                            "VALUES('" + lbl_liquidacion.Text + "', '" + lblOrdenCompra.Text + "', '" + cantidad.ToString() + "', '" + DateTime.Now.ToShortDateString() + "', 'EXPORTACION')";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        //ACTUALIZAR ORDEN DE COMPRA CAMPO liquidacion
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '" + lbl_liquidacion.Text + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET surtido_oc = surtido_oc + '" + cantidad.ToString() + "', unidad_oc = '" + lbl_liquidacion.Text + "' " +
                            "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        cmnd1.ExecuteNonQuery();
                        cmnd1.Dispose();

                        ////if (chkRecalculo.Checked == true)
                        ////{
                        //    //RECALCULO DE ORDEN DE COMPRA
                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND conse <> '" + lblConse.Text + "'";
                        //    decimal importe_det = 0;
                        //    reader1 = cmnd1.ExecuteReader();
                        //    if (reader1.HasRows)
                        //    {
                        //        reader1.Read();
                        //        importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                        //    }
                        //    reader1.Close();
                        //    reader1.Dispose();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                        //        "subtotal_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                        //        "total_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();

                        //    cmnd1 = thisConnection.CreateCommand();
                        //    cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                        //        "cantidad_oc = '" + var_dec_unidades + "', " +
                        //        "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                        //        "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                        //        "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "' AND conse = '" + lblConse.Text + "'";
                        //    cmnd1.ExecuteNonQuery();
                        //    cmnd1.Dispose();
                        ////}
                    }
                }
                //ASIGNACION DE ORDEN DE COMPRA ANTICIPADA A LIQUIDACION 
                //28/07/2021

                thisConnection.Close();

                afecta_notas_credito_exportacion(lbl_liquidacion.Text, lbl_fecha1.Text, lbl_fecha2.Text, lbl_cveprod.Text, lbl_cveprov.Text);

                MessageBox.Show("Datos Guardados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //printDocument1.Print();
                }

                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                ////Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void modificarnacional()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;
            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");
                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }

            string cvep = "";
            string var_chr_prod_clave = "";

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();
            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();

                query = "UPDATE tb_mstr_liquidacion SET uni_nac = '" + var_dec_unidades + "', liq_pre_uni = '" + var_dec_precio + "', liq_porcen1 = '" + var_dec_prod_comision + "', liq_porcen2 = '" + var_dec_prod_comision2 + "', " +
                    "liq_imp_tot = '" + Convert.ToDecimal(txt_total.Text).ToString("0.000") + "', liq_imp_por = '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.000") + "', liq_imp_liq = '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.000") + "', " +
                    "liq_costo1 = '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.000") + "', liq_costo2 = '" + ((txt_nuevocosto.Text == "") ? Convert.ToDecimal(txt_nuevocosto.Text).ToString("0.00") : "0") + "', " +
                    "liq_por_des = '" + Convert.ToDecimal(txt_porce_desc.Text).ToString("0.000") + "', liq_imp_pordes = '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.000") + "', status = 'A' " +
                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                //ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Modificación de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "M", "4.1", lbl_liquidacion.Text, "MODIFICACION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //----------14/09/2017----------//
                //ANTES DE BORRAR DE LA TABLA (tb_det_liquidacion) CONSULTAR LOS PRESTAMOS (tb_det_prestamo) PARA REGRESAR LAS CANTIDADES
                //RESTAR LA CANTIDAD ORIGINAL EN Tb_Prestamos_Prov[Saldo] - tb_det_prestamo[cantidad]
                foreach (DataRow rA in tcon.Select("valor <> ''"))
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT cantidad, tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + rA["valor"].ToString() + "' " +
                        "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N' AND estatus = 'A' AND tipo_mov = 'LQ'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "UPDATE tb_det_prestamo SET estatus = 'C' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "' " +
                                "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N' AND estatus = 'A' AND tipo_mov = 'LQ'";
                            cmnd2.ExecuteNonQuery();
                            cmnd2.Dispose();

                            if (Convert.ToDecimal(reader1["tipo_cambio"].ToString().Trim()) > 1)//ES DE DOLARES
                            {
                                string cantidade = reader1["cantidad"].ToString().Trim();
                                //CONVERTIR A DOLARES CANTIDADE
                                decimal cantidades = Math.Round(Convert.ToDecimal(cantidade) / Convert.ToDecimal(lblTipoCambio.Text), 4);
                                //string CANTIDADE2 = Convert.ToDecimal()
                                cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidades + "' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "'";
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                            }
                            else//ES DE PESOS
                            {
                                string cantidade = reader1["cantidad"].ToString().Trim();
                                cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade + "' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "'";
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                            }
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                }
                //FIN ANTES DE BORRAR DE LA TABLA (tb_det_liquidacion) CONSULTAR LOS PRESTAMOS (tb_det_prestamo) PARA REGRESAR LAS CANTIDADES
                //FIN RESTAR LA CANTIDAD ORIGINAL EN Tb_Prestamos_Prov[Saldo] - tb_det_prestamo[cantidad]
                //----------FIN 14/09/2017----------//



                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND tipo_con = 'N'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();

                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();

                    calc = tcon.Rows[i]["calculo"].ToString();

                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + lblIdPrestamo.Text + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido - '" + lblCantPrestamo.Text + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //    //    " AND Lin_Clave = '" + txt_lincve.Text + "' AND Prov_Clave = '" + lbl_cveprov.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //    //    " AND Lin_Clave = '" + txt_lincve.Text + "' AND Prov_Clave = '" + lbl_cveprov.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{
                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'N', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}
                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }

                        //----------30/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION NACIONAL
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "DOLARES")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'N', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) / Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'N', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }


                        }
                        //---------- FIN 30/11/2017----------//

                        //----------15/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) == 0)
                        //        continue;
                        //    query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', 'A', 'N', 'LQ')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 15/09/2017----------//

                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE tb_det_prestamo SET cantidad = '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //        " AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }

                }


                cuerpo += "</table>";
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);
                //printDocument1.Print();

                //16/08/2021 recalcular orden de compra por modificacion
                if (lblOrdenCompra.Text != "-")
                {
                    string összeg = var_dec_unidades.ToString();//cantidad
                    string termék = lbl_cveprod.Text;
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "UPDATE tb_det_anticipada_pt SET liq_cantidad = '" + összeg + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND liq_folio = '" + lbl_liquidacion.Text + "' " +
                        "AND liq_tipo = 'NACIONAL'";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    ////RECALCULO DE ORDEN DE COMPRA
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave <> '" + termék + "'";
                    //decimal importe_det = 0;
                    //reader1 = cmnd1.ExecuteReader();
                    //if (reader1.HasRows)
                    //{
                    //    reader1.Read();
                    //    importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                    //}
                    //reader1.Close();
                    //reader1.Dispose();
                    //cmnd1.Dispose();

                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                    //    "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                    //    "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                    //    "subtotal_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                    //    "total_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                    //    "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();

                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                    //    "cantidad_oc = '" + összeg + "', " +
                    //    "surtido_oc = '" + összeg + "', " +
                    //    "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                    //    "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                    //    "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + termék + "' AND unidad_oc = '" + lbl_liquidacion.Text + "'";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();


                }

                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }


                btnGuarda.Enabled = true;
                thisConnection.Close();
                MessageBox.Show("Datos modificados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //btnImprime_Click(null, null);
                }

                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                //////Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void modificarexportacion()
        {
            if (Convert.ToDecimal(lbl_cajas.Text) <= 0)
            {
                MessageBox.Show("El valor de cajas por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Convert.ToDecimal(lbl_flejes.Text) <= 0)
            {
                MessageBox.Show("El valor de flejes por palet debe ser mayor a 0", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Convert.ToDecimal(txt_total.Text) <= 0 || Convert.ToDecimal(txt_liquidar.Text) <= 0 || Convert.ToDecimal(txt_costounitario.Text) <= 0)
            {
                MessageBox.Show("Los importes son menores a 0 o las cantidades no son correctas, verifique por favor", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (tcon.Rows.Count == 0)
            {
                MessageBox.Show("No hay conceptos de liquidación", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tcon.Clear();
            DataRow rr;
            for (int i = 0; i < dtgConceptos.Rows.Count; i++)
            {
                rr = tcon.NewRow();
                rr["cve_con"] = dtgConceptos.Rows[i].Cells[0].Value.ToString();
                rr["nombre_con"] = dtgConceptos.Rows[i].Cells[1].Value.ToString();
                rr["unidades"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[2].Value.ToString()).ToString("0.0000");
                rr["precio"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[3].Value.ToString()).ToString("0.0000");
                rr["total"] = Convert.ToDecimal(dtgConceptos.Rows[i].Cells[4].Value.ToString()).ToString("0.0000");
                rr["valor"] = dtgConceptos.Rows[i].Cells["valor"].Value.ToString();
                rr["moni"] = (dtgConceptos.Rows[i].Cells["mon"].Value != null) ? dtgConceptos.Rows[i].Cells["mon"].Value.ToString() : "";
                rr["conse"] = (i + 1).ToString();
                rr["calculo"] = (dtgConceptos.Rows[i].Cells["val"].Value != null) ? dtgConceptos.Rows[i].Cells["val"].Value.ToString() : "0";
                tcon.Rows.Add(rr);
            }

            DataTable dtAnti = new DataTable();
            dtAnti.Columns.Add("movi", typeof(string));
            foreach (DataRow y in tcon.Select("valor <> '' AND precio > 0"))
            {
                DataRow rt = dtAnti.NewRow();
                rt["movi"] = y["valor"].ToString();
                dtAnti.Rows.Add(rt);
            }

            string cvep = "";
            string var_chr_prod_clave = "";
            decimal canti = 0;

            cvep = lbl_cveprod.Text;
            var_chr_prod_clave = lbl_cveprod.Text;
            canti = 0;

            if (var_chr_prod_clave == "05005LETOR" || var_chr_prod_clave == "05005LETOT")
            {
                if (tcon.Rows[0][0].ToString() == "1")
                {
                    canti = Math.Round(Convert.ToDecimal(lbl_libras.Text) * Convert.ToDecimal(tcon.Rows[0]["unidades"].ToString()), 2);
                }
            }

            if (MessageBox.Show("Desea realizar algún cambio en la Liquidación", "SISEMP", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.Cancel;
                return;
            }
            else
                this.DialogResult = DialogResult.OK;

            btnGuarda.Enabled = false;
            string query = "";
            thisConnection.Open();
            try
            {
                string var_dec_precio = "";
                string var_dec_prod_comision = "";
                string var_dec_prod_comision2 = "";
                string var_dec_unidades = "";

                var_dec_prod_comision = txt_valor_por.Text;
                var_dec_prod_comision2 = "0";
                var_dec_unidades = tcon.Rows[0]["unidades"].ToString();
                var_dec_precio = tcon.Rows[0]["precio"].ToString();

                query = "UPDATE tb_mstr_liquidacion SET uni_exp = '" + var_dec_unidades + "', liq_preunie = '" + var_dec_precio + "', liq_porcen_1e = '" + var_dec_prod_comision + "', liq_porcen_2e = '" + var_dec_prod_comision2 + "', " +
                    "liq_imp_tote = '" + Convert.ToDecimal(txt_total.Text).ToString("0.000") + "', liq_imp_pore = '" + Convert.ToDecimal(txt_porcentaje.Text).ToString("0.000") + "', liq_imp_liqe = '" + Convert.ToDecimal(txt_liquidar.Text).ToString("0.000") + "', " +
                    "liq_costo1e = '" + Convert.ToDecimal(txt_costounitario.Text).ToString("0.000") + "', liq_costo2e = '" + ((txt_nuevocosto.Text == "") ? Convert.ToDecimal(txt_nuevocosto.Text).ToString("0.00") : "0") + "', " +
                    "liq_exp_pordes = '" + Convert.ToDecimal(txt_porce_desc.Text).ToString("0.000") + "', liq_exp_imppordes = '" + Convert.ToDecimal(txt_cant_porce.Text).ToString("0.000") + "', status = 'A', liq_precambio = '" + Convert.ToDecimal(txt_tipocambio.Text) + "' " +
                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = query;
                cmnd1.ExecuteNonQuery();
                //ultimo_folio = Convert.ToString(cmnd1.ExecuteScalar()).Trim();
                cmnd1.Dispose();

                string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                using (StreamWriter sw = File.AppendText(filelog))
                {
                    sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Inserción de liquidación: " + lbl_liquidacion.Text);
                    sw.Close();
                }

                Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "A", "4.1", lbl_liquidacion.Text, "INSERCION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                //----------14/09/2017----------//
                //ANTES DE BORRAR DE LA TABLA (tb_det_liquidacion) CONSULTAR LOS PRESTAMOS (tb_det_prestamo) PARA REGRESAR LAS CANTIDADES
                //RESTAR LA CANTIDAD ORIGINAL EN Tb_Prestamos_Prov[Saldo] - tb_det_prestamo[cantidad]
                foreach (DataRow rA in tcon.Select("valor <> ''"))
                {
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT cantidad, tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + rA["valor"].ToString() + "' " +
                        "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E' AND estatus = 'A' AND tipo_mov = 'LQ'";
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            cmnd2 = thisConnection.CreateCommand();
                            cmnd2.CommandText = "UPDATE tb_det_prestamo SET estatus = 'C' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "' " +
                                "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E' AND estatus = 'A' AND tipo_mov = 'LQ'";
                            cmnd2.ExecuteNonQuery();
                            cmnd2.Dispose();



                            if (Convert.ToDecimal(reader1["tipo_cambio"].ToString().Trim()) > 1) //liquidacion en dolares pero prestamo en pesos
                            {
                                string cantidade = reader1["cantidad"].ToString().Trim();
                                decimal cantidades = Math.Round(Convert.ToDecimal(cantidade) * Convert.ToDecimal(txtTipoCambioResp.Text), 4);
                                cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidades + "' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "'";
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                            }
                            else
                            {
                                string cantidade = reader1["cantidad"].ToString().Trim();

                                cmnd2 = thisConnection.CreateCommand();
                                cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade + "' WHERE Id_Movimiento = '" + rA["valor"].ToString() + "'";
                                cmnd2.ExecuteNonQuery();
                                cmnd2.Dispose();
                            }


                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                }
                //FIN ANTES DE BORRAR DE LA TABLA (tb_det_liquidacion) CONSULTAR LOS PRESTAMOS (tb_det_prestamo) PARA REGRESAR LAS CANTIDADES
                //FIN RESTAR LA CANTIDAD ORIGINAL EN Tb_Prestamos_Prov[Saldo] - tb_det_prestamo[cantidad]
                //----------FIN 14/09/2017----------//

                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND tipo_con = 'E'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();



                //Detalles de la liquidacion
                string cvex = "";
                string nomx = "";
                decimal unix = 0;
                decimal prex = 0;
                decimal totx = 0;
                string valx = "";
                string conx = "";

                string calc = "";

                string cuerpo = "<table>";

                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    cvex = tcon.Rows[i][0].ToString();
                    nomx = tcon.Rows[i][1].ToString().Replace("'", " ");
                    unix = Convert.ToDecimal(tcon.Rows[i][2].ToString());
                    prex = Convert.ToDecimal(tcon.Rows[i][3].ToString());
                    totx = Convert.ToDecimal(tcon.Rows[i][4].ToString());
                    valx = tcon.Rows[i]["valor"].ToString();
                    conx = tcon.Rows[i]["conse"].ToString();

                    calc = tcon.Rows[i]["calculo"].ToString();

                    try
                    {
                        //if (cvex == "95")
                        //{
                        //    query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + lblIdPrestamo.Text + "')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido - '" + lblCantPrestamo.Text + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //    //    " AND Lin_Clave = '" + txt_lincve.Text + "' AND Prov_Clave = '" + lbl_cveprov.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();

                        //    //query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //    //    " AND Lin_Clave = '" + txt_lincve.Text + "' AND Prov_Clave = '" + lbl_cveprov.Text + "'";
                        //    //cmnd2 = thisConnection.CreateCommand();
                        //    //cmnd2.CommandText = query;
                        //    //cmnd2.ExecuteNonQuery();
                        //    //cmnd2.Dispose();
                        //}
                        //else
                        //{
                        query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con, Id_Movimiento, conse, calculo) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E', '" + valx + "', '" + conx + "', '" + calc + "')";
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = query;
                        cmnd2.ExecuteNonQuery();
                        cmnd2.Dispose();
                        //}
                        if (valx != "")
                        {
                            cuerpo += "<tr>" + query + "</tr>";
                        }
                        //----------30/11/2017----------//
                        //SE GUARDAN LOS DATOS PARA LA LIQUIDACION EXPORTACION
                        if (tcon.Rows[i]["valor"].ToString() != "")
                        {
                            if (prex > 0)
                            {
                                if (tcon.Rows[i]["moni"].ToString() == "PESOS")
                                {
                                    //PREX ESTA EN PESOS
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', " +
                                        "'A', 'E', 'LQ', '" + lblTipoCambio.Text + "', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    //ORDEN DE COMPRA NACIONAL
                                    //SI EL ANTICIPO ES EN DOLARES HACER LA CONVERSION DE PESOS A DOLARES PARA EL INCREMENTO DEL SALDO
                                    string prexA = Convert.ToDecimal(Convert.ToDecimal(prex) * Convert.ToDecimal(lblTipoCambio.Text)).ToString("0.0000");

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prexA + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                else
                                {
                                    cmnd1.CommandText = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov, tipo_cambio, fecha_mov) " +
                                        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex + "', " +
                                        "'A', 'E', 'LQ', '1', '" + DateTime.Now.ToShortDateString() + "')";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();

                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                            }
                        }
                        //---------- FIN 30/11/2017----------//

                        //----------13/09/2017----------//
                        //GUARDADO DE DATOS EN tb_det_prestamos
                        //if (tcon.Rows[i]["valor"].ToString() != "")
                        //{
                        //    if (Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) == 0)
                        //        continue;
                        //    query = "INSERT INTO tb_det_prestamo(Id_Movimiento, liq_folio, cantidad, estatus, liq_tipo, tipo_mov) " +
                        //        "VALUES('" + tcon.Rows[i]["valor"].ToString() + "', '" + lbl_liquidacion.Text + "', '" + prex.ToString("0.0000") + "', 'A', 'E', 'LQ')";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();

                        //    query = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo + '" + prex.ToString("0.0000") + "' WHERE Id_Movimiento = '" + tcon.Rows[i]["valor"].ToString() + "'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //    cmnd2.Dispose();
                        //}
                        //FIN GUARDADO DE DATOS EN tb_det_prestamos
                        //----------FIN 13/09/2017----------//
                        //query = "INSERT INTO tb_det_liquidacion(liq_folio, cve_con, nom_con, unid_con, precio_con, importe_con, tipo_con) VALUES('" + lbl_liquidacion.Text + "', '" + cvex + "'," +
                        //    " '" + nomx + "', " + unix.ToString("0.0000") + ", " + prex.ToString("0.0000") + ", " + totx.ToString("0.0000") + ", 'E')";
                        //cmnd2 = thisConnection.CreateCommand();
                        //cmnd2.CommandText = query;
                        //cmnd2.ExecuteNonQuery();
                        //cmnd2.Dispose();

                        //if (cvex == "95")
                        //{
                        //    query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido - '" + lblCantPrestamo.Text + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //        " AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();

                        //    query = "UPDATE Tb_Prestamos_Prov SET surtido = surtido + '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //        " AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();

                        //    query = "UPDATE tb_det_prestamo SET cantidad = '" + totx.ToString("0.0000") + "' WHERE Id_Movimiento = '" + lblIdPrestamo.Text + "'" +
                        //        " AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E'";
                        //    cmnd2 = thisConnection.CreateCommand();
                        //    cmnd2.CommandText = query;
                        //    cmnd2.ExecuteNonQuery();
                        //}
                    }
                    catch (SqlException sqlex)
                    {
                        MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (thisConnection.State == ConnectionState.Open)
                            thisConnection.Close();
                        Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                        Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                        //this.Close();                                
                        return;
                    }
                }

                cuerpo += "</table>";
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", cuerpo);
                //printDocument1.Print();

                //16/08/2021 recalcular orden de compra por modificacion
                if (lblOrdenCompra.Text != "-")
                {
                    string összeg = var_dec_unidades.ToString();//cantidad
                    string termék = lbl_cveprod.Text;
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "UPDATE tb_det_anticipada_pt SET liq_cantidad = '" + összeg + "' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND liq_folio = '" + lbl_liquidacion.Text + "' " +
                        "AND liq_tipo = 'EXPORTACION'";
                    cmnd1.ExecuteNonQuery();
                    cmnd1.Dispose();

                    ////RECALCULO DE ORDEN DE COMPRA
                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "SELECT ISNULL(SUM(importe_oc), 0) AS importe_det FROM tb_det_ordenescompra WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave <> '" + termék + "'";
                    //decimal importe_det = 0;
                    //reader1 = cmnd1.ExecuteReader();
                    //if (reader1.HasRows)
                    //{
                    //    reader1.Read();
                    //    importe_det = Convert.ToDecimal(reader1["importe_det"].ToString().Trim());
                    //}
                    //reader1.Close();
                    //reader1.Dispose();
                    //cmnd1.Dispose();

                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET " +
                    //    "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                    //    "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "', " +
                    //    "subtotal_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "', " +
                    //    "total_oc = '" + (Convert.ToDecimal(txt_liquidar.Text) + importe_det).ToString() + "' " +
                    //    "WHERE numero_oc = '" + lblOrdenCompra.Text + "'";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();

                    //cmnd1 = thisConnection.CreateCommand();
                    //cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET " +
                    //    "cantidad_oc = '" + összeg + "', " +
                    //    "surtido_oc = '" + összeg + "', " +
                    //    "precio_oc = '" + Convert.ToDecimal(txt_costounitario.Text).ToString() + "', " +
                    //    "importe_oc = '" + Convert.ToDecimal(txt_liquidar.Text).ToString() + "' " +
                    //    "WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + termék + "' AND unidad_oc = '" + lbl_liquidacion.Text + "'";
                    //cmnd1.ExecuteNonQuery();
                    //cmnd1.Dispose();


                }


                if (dtAnti.Rows.Count > 0)
                {
                    string cad = correo_movimientos(dtAnti);
                    enviarcorreo(cad);
                }

                btnGuarda.Enabled = true;
                thisConnection.Close();
                MessageBox.Show("Datos Modificados", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (lbl_cveprov.Text != "03")// Modificado 31/01/2024
                {
                    //btnImprime_Click(null, null);
                }

                //guardar en servidor
                printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
                printDocument1.Print();

                FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

                FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                if (liq_copy.Exists == true)
                {
                    liq_copy.Delete();
                }
                archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                //////Process.Start(@"\\gabira1\liquidaciones\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            }
            catch (SqlException sqlex)
            {
                MessageBox.Show("Error de sistema, no se termino de guardar la liquidacion", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (thisConnection.State == ConnectionState.Open)
                    thisConnection.Close();
                Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "4.1", lbl_liquidacion.Text + " " + sqlex.ToString() + " " + query, "SISEMP");
                Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + sqlex.ToString());
                //this.Close();                                
                return;
            }
        }

        private void textoliquidacion()
        {
            TextWriter tw = new StreamWriter("c:\\empaque\\Gab\\liquidacion1.txt");

            tw.Write("/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*------LIQUIDACION NO VALIDA------/*/*/*/*/*/*/*/*/*/*/*/*/*/*"); tw.WriteLine();
            tw.Write("Pag: 1");
            tw.Write("\t\t\t Comercializadora GAB, S.A. de C.V."); tw.WriteLine();
            tw.Write("Liquidacion de Producto Terminado del :" + lbl_fecha1.Text + " al " + lbl_fecha2.Text); tw.WriteLine();
            tw.Write("Proveedor: " + lbl_cveprov.Text + " " + lbl_proveedor.Text + "\t\t Liquidacion: SIN FOLIO"); tw.WriteLine();
            tw.Write("Producto: " + lbl_cveprod.Text + " " + lbl_producto.Text); tw.WriteLine();
            if (procedencia == "NACIONAL")
            {
                tw.Write("Nacional" + "\t\t\t\t\t" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString()); tw.WriteLine();
            }
            if (procedencia == "EXPORTACION")
            {
                tw.Write("Exportacion " + ((txt_tipocambio.Text != "") ? "Tipo cambio " + txt_tipocambio.Text + "\t\t\t" : "\t\t\t\t\t") + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString()); tw.WriteLine();
            }
            tw.Write("\t\t\t Flejes x pallet: " + lbl_flejes.Text + "\t Cajas por pallet: " + lbl_cajas.Text); tw.WriteLine();
            tw.Write("\t\t\t\t\t Total \t\t Precio \t\t Importe \t\t Importe"); tw.WriteLine();
            tw.Write("\t\t\t\t\t Unidades \t\t Unitario \t\t Total \t\t x Caja"); tw.WriteLine();

            decimal totcaj = 0;
            decimal totcaj2 = 0;
            decimal totalx = 0;
            for (int x = 0; x < dtgConceptos.Rows.Count; x++)
            {
                totcaj = Convert.ToDecimal(dtgConceptos.Rows[0].Cells[2].Value.ToString());
                totcaj2 = Convert.ToDecimal(dtgConceptos.Rows[x].Cells[2].Value.ToString());

                if (dtgConceptos.Rows[x].Cells[5].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgConceptos.Rows[x].Cells[3].Value.ToString()) == 0)
                        continue;
                }

                for (int y = 1; y < dtgConceptos.Columns.Count; y++)
                {
                    if (y == 1)
                    {
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length > 25)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().Substring(0, 25));
                        }
                        else
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value);
                        }

                    }
                    if (y == 2)
                    {

                        if (y != dtgConceptos.Columns.Count - 1)
                        {
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                        }
                    }
                    if (y == 3)
                    {
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                    }
                    if (y == 4)
                    {
                        decimal aa = 0;
                        if (dtgConceptos.Rows[x].Cells[0].Value.ToString().Length > 2)
                        {
                            string val2 = verificarunidad2(dtgConceptos.Rows[x].Cells[0].Value.ToString());
                            if (val2 == "ROL" || val2 == "ROLLO")
                            {
                                aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj2);
                            }
                            else
                                aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj);
                        }
                        else
                            aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj);
                        if (x > 0)
                        {
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                totalx = totalx + aa;
                            //else
                            //    totalx = totalx - aa;
                        }


                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                //decimal A = aa * -1;
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                //decimal A = aa * -1;
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                //decimal A = aa * -1;
                                decimal A = aa;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa; //decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa; //decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa; //decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }

                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa; //decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa; //decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }

                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                    }
                }
                tw.WriteLine();
            }
            tw.Write("\t\t\t\t\t\t\t\t\t" + totalx.ToString("###,##0.000").PadLeft(37, ' ')); tw.WriteLine();
            tw.Write("\r\t\t\t\t\t TOTAL" + txt_total.Text.ToString().PadLeft(36, ' ')); tw.WriteLine();
            tw.Write("\t\t\t\t\t" + lbl_porcentaje.Text + txt_porcentaje.Text.ToString().PadLeft(35, ' ')); tw.Write("\t\t" + Math.Round(Convert.ToDecimal(txt_porcentaje.Text) * -1 / Convert.ToDecimal(cantidad), 3).ToString()); tw.WriteLine();
            tw.Write(label15.Text + "\t\t% " + txt_porce_desc.Text + txt_cant_porce.Text.ToString().PadLeft(39, ' '));
            tw.Write("\t\t" + Math.Round(Convert.ToDecimal(txt_cant_porce.Text) * -1 / Convert.ToDecimal(cantidad), 3).ToString()); tw.WriteLine();
            if (procedencia == "NACIONAL")
                tw.Write("\t\t\t\t\t TOTAL A LIQUIDAR" + txt_liquidar.Text.ToString().PadLeft(25, ' ') + " M.N."); tw.WriteLine();
            if (procedencia == "EXPORTACION")
                tw.Write("\t\t\t\t\t TOTAL A LIQUIDAR" + txt_liquidar.Text.ToString().PadLeft(25, ' ') + " USD"); tw.WriteLine();
            tw.Write("\r\t\t\t\t\t COSTO UNITARIO" + txt_costounitario.Text.ToString().PadLeft(25, ' ')); tw.WriteLine();

            tw.Write("_____________________________________________________________________________________"); tw.WriteLine();
            //Copia
            totalx = 0;
            tw.Write("\rCopia");
            tw.Write("\t\t\t Comercializadora GAB, S.A. de C.V."); tw.WriteLine();
            tw.Write("Liquidacion de Producto Terminado del :" + lbl_fecha1.Text + " al " + lbl_fecha2.Text); tw.WriteLine();
            tw.Write("Proveedor: " + lbl_cveprov.Text + " " + lbl_proveedor.Text + "\t\t Liquidacion: SIN FOLIO"); tw.WriteLine();
            tw.Write("Producto: " + lbl_cveprod.Text + " " + lbl_producto.Text); tw.WriteLine();
            if (procedencia == "NACIONAL")
            {
                tw.Write("Nacional" + "\t\t\t\t\t" + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString()); tw.WriteLine();
            }
            if (procedencia == "EXPORTACION")
            {
                tw.Write("Exportacion " + ((txt_tipocambio.Text != "") ? "Tipo cambio " + txt_tipocambio.Text + "\t\t\t" : "\t\t\t\t\t") + DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString()); tw.WriteLine();
            }
            tw.Write("\t\t\t Flejes x pallet: " + lbl_flejes.Text + "\t Cajas por pallet: " + lbl_cajas.Text); tw.WriteLine();
            tw.Write("\t\t\t\t\t Total \t\t Precio \t\t Importe \t\t Importe"); tw.WriteLine();
            tw.Write("\t\t\t\t\t Unidades \t\t Unitario \t\t Total \t\t x Caja"); tw.WriteLine();

            for (int x = 0; x < dtgConceptos.Rows.Count; x++)
            {
                totcaj = Convert.ToDecimal(dtgConceptos.Rows[0].Cells[2].Value.ToString());
                totcaj2 = Convert.ToDecimal(dtgConceptos.Rows[x].Cells[2].Value.ToString());

                if (dtgConceptos.Rows[x].Cells[5].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(dtgConceptos.Rows[x].Cells[3].Value.ToString()) == 0)
                        continue;
                }

                for (int y = 1; y < dtgConceptos.Columns.Count; y++)
                {
                    if (y == 1)
                    {
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length > 25)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().Substring(0, 25));
                        }
                        else
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value);
                        }

                    }
                    if (y == 2)
                    {
                        if (y != dtgConceptos.Columns.Count - 1)
                        {
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                            if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                            {
                                int numcad = 0;
                                if (dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length > 25)
                                    numcad = (37 - 25);
                                else
                                    numcad = (37 - dtgConceptos.Rows[x].Cells[y - 1].Value.ToString().Length);
                                tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(numcad, ' '));
                            }
                        }
                    }
                    if (y == 3)
                    {
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                        }
                    }
                    if (y == 4)
                    {
                        decimal aa = 0;
                        if (dtgConceptos.Rows[x].Cells[0].Value.ToString().Length > 2)
                        {
                            string val2 = verificarunidad2(dtgConceptos.Rows[x].Cells[0].Value.ToString());
                            if (val2 == "ROL" || val2 == "ROLLO")
                            {
                                aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj2);
                            }
                            else
                                aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj);
                        }
                        else
                            aa = (Math.Abs(Convert.ToDecimal(dtgConceptos.Rows[x].Cells[y].Value.ToString())) / totcaj);
                        if (x > 0)
                        {
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                totalx = totalx + aa;
                            //else
                            //    totalx = totalx - aa;
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 15)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() == "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 14)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 13)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 12)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 11)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 10)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 9)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 8)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 7)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 6)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 5)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.00").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 4)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 3)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 2)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                        if (dtgConceptos.Rows[x].Cells[y].Value.ToString().Length == 1)
                        {
                            tw.Write(dtgConceptos.Rows[x].Cells[y].Value.ToString().PadLeft(16, ' '));
                            if (dtgConceptos.Rows[x].Cells[0].Value.ToString() != "93")
                                tw.Write(aa.ToString("###,##0.000").PadLeft(16, ' '));
                            else
                            {
                                decimal A = aa;//decimal A = aa * -1;
                                tw.Write(A.ToString("###,##0.000").PadLeft(16, ' '));
                            }
                        }
                    }
                }
                tw.WriteLine();
            }
            tw.Write("\t\t\t\t\t\t\t\t\t" + totalx.ToString("###,##0.000").PadLeft(37, ' ')); tw.WriteLine();
            tw.Write("\r\t\t\t\t\t TOTAL" + txt_total.Text.ToString().PadLeft(36, ' ')); tw.WriteLine();
            tw.Write("\t\t\t\t\t" + lbl_porcentaje.Text + txt_porcentaje.Text.ToString().PadLeft(35, ' ')); tw.Write("\t\t" + Math.Round(Convert.ToDecimal(txt_porcentaje.Text) * -1 / Convert.ToDecimal(cantidad), 3).ToString()); tw.WriteLine();
            tw.Write(label15.Text + "\t\t% " + txt_porce_desc.Text + txt_cant_porce.Text.ToString().PadLeft(39, ' ')); tw.Write("\t\t" + Math.Round(Convert.ToDecimal(txt_cant_porce.Text) * -1 / Convert.ToDecimal(cantidad), 3).ToString()); tw.WriteLine();
            if (procedencia == "NACIONAL")
                tw.Write("\t\t\t\t\t TOTAL A LIQUIDAR" + txt_liquidar.Text.ToString().PadLeft(25, ' ') + " M.N."); tw.WriteLine();
            if (procedencia == "EXPORTACION")
                tw.Write("\t\t\t\t\t TOTAL A LIQUIDAR" + txt_liquidar.Text.ToString().PadLeft(25, ' ') + " USD"); tw.WriteLine();
            tw.Write("\r\t\t\t\t\t COSTO UNITARIO" + txt_costounitario.Text.ToString().PadLeft(25, ' ')); tw.WriteLine();
            tw.Close();

            string filelog = "C:\\SisEmpWeb\\eventlog.txt";
            using (StreamWriter sw = File.AppendText(filelog))
            {
                sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Impresión de liquidacion: " + lbl_liquidacion.Text);
                sw.Close();
            }


            Process.Start("wordpad.exe", "c:\\empaque\\Gab\\liquidacion1.txt");
        }

        private void consultaliquidacionN()
        {
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT liq_cajas_pal, liq_flejes_pal, liq_unidades, liq_mermas, uni_nac, uninac_oc, uni_mern, liq_pre_uni, liq_porcen1, liq_porcen2, liq_imp_tot, liq_imp_por, liq_imp_liq, liq_costo1, liq_costo2," +
                " liq_ocompra, liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8, liq_cantiocn, liq_por_des, liq_imp_pordes, liq_nac, liq_libras FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    lbl_cajas.Text = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim()).ToString("###,##0.000");
                    lbl_flejes.Text = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim()).ToString("###,##0.000");
                    txt_valor_por.Text = Convert.ToDecimal(reader1.GetValue(8).ToString().Trim()).ToString("###,##0.000");
                    txt_total.Text = Convert.ToDecimal(reader1.GetValue(10).ToString().Trim()).ToString("###,##0.000");
                    txt_porcentaje.Text = reader1.GetValue(11).ToString().Trim();
                    txt_liquidar.Text = Convert.ToDecimal(reader1.GetValue(12).ToString().Trim()).ToString("###,##0.000");
                    txt_costounitario.Text = reader1.GetValue(13).ToString().Trim();
                    txt_porce_desc.Text = reader1.GetValue(25).ToString().Trim();
                    txt_cant_porce.Text = Convert.ToDecimal(reader1.GetValue(26).ToString().Trim()).ToString("###,##0.000");
                    lbl_libras.Text = Convert.ToDecimal(reader1.GetValue(28).ToString().Trim()).ToString("###,##0.000");
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT cve_con, nom_con, unid_con, precio_con, importe_con, Id_Movimiento, calculo FROM tb_det_liquidacion WHERE tipo_con = 'N' AND liq_folio = '" + lbl_liquidacion.Text + "' ORDER BY conse";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = tcon.NewRow();
                    rw["cve_con"] = reader1.GetValue(0).ToString().Trim();
                    rw["nombre_con"] = reader1.GetValue(1).ToString().Trim();
                    rw["unidades"] = reader1.GetValue(2).ToString().Trim();
                    rw["precio"] = reader1.GetValue(3).ToString().Trim();
                    rw["total"] = reader1.GetValue(4).ToString().Trim();
                    rw["valor"] = reader1["Id_Movimiento"].ToString().Trim();
                    rw["calculo"] = reader1["calculo"].ToString().Trim();
                    string valX = reader1["Id_Movimiento"].ToString().Trim();
                    if (valX != "")
                    {
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT Moneda, Total, Saldo FROM Tb_Prestamos_Prov WHERE Id_Movimiento = '" + valX + "'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            decimal total_pre = 0;
                            decimal saldo_pre = 0;
                            decimal saldo_total = 0;
                            total_pre = Convert.ToDecimal(reader2["Total"].ToString().Trim());
                            saldo_pre = Convert.ToDecimal(reader2["Saldo"].ToString().Trim());
                            saldo_total = total_pre - saldo_pre;
                            rw["moni"] = reader2["Moneda"].ToString().Trim();
                            rw["saldo"] = saldo_total.ToString();
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + valX + "' AND liq_tipo = 'N' AND liq_folio = '" + lbl_liquidacion.Text + "'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            rw["tc"] = reader2["tipo_cambio"].ToString().Trim();
                            if (Convert.ToDecimal(reader2["tipo_cambio"].ToString().Trim()) == 1)
                            {
                                rw["saldo"] = (Convert.ToDecimal(rw["saldo"]) + Convert.ToDecimal(rw["precio"]));
                            }
                            else
                            {
                                decimal importe_descontado = 0;
                                importe_descontado = Convert.ToDecimal(rw["precio"]);
                                rw["saldo"] = (Convert.ToDecimal(rw["saldo"]) * Convert.ToDecimal(rw["tc"])) + importe_descontado;
                            }
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();
                    }
                    tcon.Rows.Add(rw);

                    //if (reader1.GetValue(0).ToString().Trim() == "95")
                    //{
                    //    lblIdPrestamo.Text = reader1.GetValue(5).ToString().Trim();
                    //    lblCantPrestamo.Text = reader1.GetValue(4).ToString().Trim();
                    //    //cmnd2 = thisConnection.CreateCommand();
                    //    //cmnd2.CommandText = "SELECT Id_Movimiento, cantidad FROM tb_det_prestamo WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N'";
                    //    //reader2 = cmnd2.ExecuteReader();
                    //    //if (reader2.HasRows)
                    //    //{
                    //    //    while (reader2.Read())
                    //    //    {
                    //    //        lblIdPrestamo.Text = reader2.GetValue(0).ToString().Trim();
                    //    //        lblCantPrestamo.Text = reader2.GetValue(1).ToString().Trim();
                    //    //    }
                    //    //}
                    //    //reader2.Close();
                    //    //reader2.Dispose();
                    //    //cmnd2.Dispose();
                    //}
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            //28/07/2021
            //BUSCAR SI LIQUIDACION ES DE UNA ORDEN DE COMPRA ANTICIPADA
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT numero_oc FROM tb_det_anticipada_pt WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'NACIONAL'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                lblOrdenCompra.Text = reader1["numero_oc"].ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            //FIN BUSCAR SI LIQUIDACION ES DE UNA ORDEN DE COMPRA ANTICIPADA

            thisConnection.Close();

            foreach (DataRow rx in tcon.Rows)
            {
                dtgConceptos.Rows.Add(rx["cve_con"].ToString(), rx["nombre_con"].ToString(), Convert.ToDecimal(rx["unidades"].ToString()).ToString("###,###,###,##0.000"),
                    Convert.ToDecimal(rx["precio"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rx["total"].ToString()).ToString("###,###,###,##0.000"),
                    rx["valor"].ToString(), rx["moni"].ToString(), rx["tc"].ToString(), rx["saldo"].ToString(), rx["calculo"].ToString());
            }

            lbl_porcentaje.Text = "% " + Convert.ToDecimal(txt_valor_por.Text).ToString("0.00");



        }

        private void consultaliquidacionE()
        {
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT liq_cajas_pal, liq_flejes_pal, liq_unidades, liq_mermas, uni_exp, uniexp_oc, uni_mere, liq_precambio, liq_preunie, liq_porcen_1e, liq_porcen_2e, liq_imp_tote, liq_imp_pore, liq_imp_liqe, liq_costo1e, liq_costo2e," +
                " liq_ocompra, liq_numoc1, liq_numoc2, liq_numoc3, liq_numoc4, liq_numoc5, liq_numoc6, liq_numoc7, liq_numoc8, liq_cantioce, liq_exp_pordes, liq_exp_imppordes, liq_exp, liq_libras FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    lbl_cajas.Text = Convert.ToDecimal(reader1.GetValue(0).ToString().Trim()).ToString("###,##0.000");
                    lbl_flejes.Text = Convert.ToDecimal(reader1.GetValue(1).ToString().Trim()).ToString("###,##0.000");
                    txt_tipocambio.Text = reader1.GetValue(7).ToString().Trim();
                    txtTipoCambioResp.Text = reader1.GetValue(7).ToString().Trim();
                    txt_valor_por.Text = Convert.ToDecimal(reader1.GetValue(9).ToString().Trim()).ToString("###,##0.000");
                    txt_total.Text = Convert.ToDecimal(reader1.GetValue(11).ToString().Trim()).ToString("###,##0.000");
                    txt_porcentaje.Text = Convert.ToDecimal(reader1.GetValue(12).ToString().Trim()).ToString("###,##0.000");
                    txt_liquidar.Text = Convert.ToDecimal(reader1.GetValue(13).ToString().Trim()).ToString("###,##0.000");
                    txt_costounitario.Text = Convert.ToDecimal(reader1.GetValue(14).ToString().Trim()).ToString("###,##0.000");
                    txt_porce_desc.Text = Convert.ToDecimal(reader1.GetValue(26).ToString().Trim()).ToString("###,##0.000");
                    txt_cant_porce.Text = Convert.ToDecimal(reader1.GetValue(27).ToString().Trim()).ToString("###,##0.000");
                    lbl_libras.Text = Convert.ToDecimal(reader1.GetValue(29).ToString().Trim()).ToString("###,##0.000");
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT cve_con, nom_con, unid_con, precio_con, importe_con, Id_Movimiento, calculo FROM tb_det_liquidacion WHERE tipo_con = 'E' AND liq_folio = '" + lbl_liquidacion.Text + "' ORDER BY conse";
            reader1 = cmnd1.ExecuteReader();
            DataRow rw;
            if (reader1.HasRows)
            {
                while (reader1.Read())
                {
                    rw = tcon.NewRow();
                    rw["cve_con"] = reader1.GetValue(0).ToString().Trim();
                    rw["nombre_con"] = reader1.GetValue(1).ToString().Trim();
                    rw["unidades"] = reader1.GetValue(2).ToString().Trim();
                    rw["precio"] = reader1.GetValue(3).ToString().Trim();
                    rw["total"] = reader1.GetValue(4).ToString().Trim();
                    rw["valor"] = reader1["Id_Movimiento"].ToString().Trim();
                    rw["calculo"] = reader1["calculo"].ToString().Trim();
                    string valX = reader1["Id_Movimiento"].ToString().Trim();
                    if (valX != "")
                    {
                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT Moneda, Total, Saldo FROM Tb_Prestamos_Prov WHERE Id_Movimiento = '" + valX + "'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            decimal total_pre = 0;
                            decimal saldo_pre = 0;
                            decimal saldo_total = 0;
                            total_pre = Convert.ToDecimal(reader2["Total"].ToString().Trim());
                            saldo_pre = Convert.ToDecimal(reader2["Saldo"].ToString().Trim());
                            saldo_total = total_pre - saldo_pre;
                            rw["moni"] = reader2["Moneda"].ToString().Trim();
                            rw["saldo"] = saldo_total.ToString();
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();

                        cmnd2 = thisConnection.CreateCommand();
                        cmnd2.CommandText = "SELECT tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + valX + "' AND liq_tipo = 'E' AND liq_folio = '" + lbl_liquidacion.Text + "'";
                        reader2 = cmnd2.ExecuteReader();
                        if (reader2.HasRows)
                        {
                            reader2.Read();
                            rw["tc"] = reader2["tipo_cambio"].ToString().Trim();
                            if (Convert.ToDecimal(reader2["tipo_cambio"].ToString().Trim()) == 1)
                            {
                                rw["saldo"] = (Convert.ToDecimal(rw["saldo"]) + Convert.ToDecimal(rw["precio"]));
                            }
                            else
                            {
                                decimal importe_descontado = 0;
                                importe_descontado = Convert.ToDecimal(rw["precio"]);
                                rw["saldo"] = (Convert.ToDecimal(rw["saldo"]) / Convert.ToDecimal(rw["tc"])) + importe_descontado;
                            }
                        }
                        reader2.Close();
                        reader2.Dispose();
                        cmnd2.Dispose();
                    }
                    tcon.Rows.Add(rw);

                    //if (reader1.GetValue(0).ToString().Trim() == "95")
                    //{
                    //    lblIdPrestamo.Text = reader1.GetValue(5).ToString().Trim();
                    //    lblCantPrestamo.Text = reader1.GetValue(4).ToString().Trim();
                    //    //cmnd2 = thisConnection.CreateCommand();
                    //    //cmnd2.CommandText = "SELECT Id_Movimiento, cantidad FROM tb_det_prestamo WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E'";
                    //    //reader2 = cmnd2.ExecuteReader();
                    //    //if (reader2.HasRows)
                    //    //{
                    //    //    while (reader2.Read())
                    //    //    {
                    //    //        lblIdPrestamo.Text = reader2.GetValue(0).ToString().Trim();
                    //    //        lblCantPrestamo.Text = reader2.GetValue(1).ToString().Trim();
                    //    //    }
                    //    //}
                    //    //reader2.Close();
                    //    //reader2.Dispose();
                    //    //cmnd2.Dispose();
                    //}
                }
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();

            //28/07/2021
            //BUSCAR SI LIQUIDACION ES DE UNA ORDEN DE COMPRA ANTICIPADA
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT numero_oc FROM tb_det_anticipada_pt WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'EXPORTACION'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                lblOrdenCompra.Text = reader1["numero_oc"].ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            //FIN BUSCAR SI LIQUIDACION ES DE UNA ORDEN DE COMPRA ANTICIPADA

            thisConnection.Close();

            foreach (DataRow rx in tcon.Rows)
            {
                dtgConceptos.Rows.Add(rx["cve_con"].ToString(), rx["nombre_con"].ToString(), Convert.ToDecimal(rx["unidades"].ToString()).ToString("###,###,###,##0.000"),
                    Convert.ToDecimal(rx["precio"].ToString()).ToString("###,###,##0.000"), Convert.ToDecimal(rx["total"].ToString()).ToString("###,###,###,##0.000"),
                    rx["valor"].ToString(), rx["moni"].ToString(), rx["tc"].ToString(), rx["saldo"], rx["calculo"].ToString());
            }

            lbl_porcentaje.Text = "% " + Convert.ToDecimal(txt_valor_por.Text).ToString("0.00");
        }

        private string verificarunidad2(string clave)
        {
            string val = "";
            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();
            cmnd1.CommandText = "SELECT um_clave FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
            reader1 = cmnd1.ExecuteReader();
            if (reader1.HasRows)
            {
                reader1.Read();
                val = reader1.GetValue(0).ToString().Trim();
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();
            return val;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.PageUnit = GraphicsUnit.Millimeter;
            Graphics g = e.Graphics;
            SolidBrush brsh = new SolidBrush(Color.Black);
            Font fuente = new System.Drawing.Font("Courier New", 10);
            g.DrawString("Pag: 1", fuente, brsh, 3, 2);
            g.DrawString("Comercializadora GAB, S.A. de C.V.", fuente, brsh, 48, 2);
            g.DrawString("Liquidación de Producto Terminado del " + lbl_fecha1.Text + " al " + lbl_fecha2.Text, fuente, brsh, 3, 6);
            g.DrawString("Proveedor:" + lbl_cveprov.Text + " " + lbl_proveedor.Text, fuente, brsh, 3, 10);
            g.DrawString("Liquidación:" + lbl_liquidacion.Text, fuente, brsh, 118, 10);
            g.DrawString("Producto:" + lbl_cveprod.Text + " " + lbl_producto.Text, fuente, brsh, 3, 14);
            if (txt_tipo.Text == "NACIONAL")
                g.DrawString("Nacional", fuente, brsh, 3, 18);
            else
                g.DrawString("Exportación" + "Tipo cambio: " + txt_tipocambio.Text, fuente, brsh, 3, 18);
            g.DrawString(DateTime.Now.ToLongDateString(), fuente, brsh, 118, 18);
            g.DrawString("Flejes x pallet: " + lbl_flejes.Text, fuente, brsh, 48, 22);
            g.DrawString("Cajas x pallet: " + lbl_cajas.Text, fuente, brsh, 118, 22);
            g.DrawString("Total", fuente, brsh, 85, 26);
            g.DrawString("Unidades", fuente, brsh, 82, 30);
            g.DrawString("Precio", fuente, brsh, 113, 26);
            g.DrawString("Unitario", fuente, brsh, 111, 30);
            g.DrawString("Importe", fuente, brsh, 153, 26);
            g.DrawString("Total", fuente, brsh, 154, 30);
            g.DrawString("Importe", fuente, brsh, 188, 26);
            g.DrawString("x caja", fuente, brsh, 189, 30);

            int ren = 34;
            decimal totalx = 0;
            int z = 0;
            decimal totcaj = 0;
            decimal totcaj2 = 0;
            foreach (DataGridViewRow rw in dtgConceptos.Rows)
            {
                totcaj = Convert.ToDecimal(dtgConceptos.Rows[0].Cells[2].Value.ToString());
                totcaj2 = Convert.ToDecimal(rw.Cells[2].Value.ToString());

                if (rw.Cells[5].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(rw.Cells[3].Value.ToString()) == 0)
                        continue;
                }

                if (rw.Cells[1].Value.ToString().Trim().Length > 34)
                    g.DrawString(rw.Cells[1].Value.ToString().Substring(0, 34).ToString(), fuente, brsh, 3, ren);
                else
                    g.DrawString(rw.Cells[1].Value.ToString(), fuente, brsh, 3, ren);

                g.DrawString(rw.Cells[2].Value.ToString(), fuente, brsh, posicionunidades(rw.Cells[2].Value.ToString()), ren);
                g.DrawString(rw.Cells[3].Value.ToString(), fuente, brsh, posicionprecio(rw.Cells[3].Value.ToString()), ren);
                g.DrawString(rw.Cells[4].Value.ToString(), fuente, brsh, posicionimporte(rw.Cells[4].Value.ToString()), ren);
                if (Convert.ToDecimal(rw.Cells[4].Value.ToString()) == 0)
                    g.DrawString("0.000", fuente, brsh, posicioncaja("0.000"), ren);
                else
                {
                    //decimal aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / Convert.ToDecimal(totcaj));
                    decimal aa = 0;
                    if (rw.Cells[0].Value.ToString().Length > 2)
                    {
                        string val2 = verificarunidad2(rw.Cells[0].Value.ToString());
                        if (val2 == "ROL" || val2 == "ROLLO")
                        {
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / totcaj2);
                        }
                        else
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / totcaj);
                    }
                    else
                    {
                        if (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) == 0)
                            aa = 0;
                        else
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / Convert.ToDecimal(totcaj));

                    }

                    if (rw.Cells[1].Value.ToString().Contains("Notas de Cargo") || rw.Cells[1].Value.ToString().Contains("NOTAS DE CARGO"))
                    {
                        g.DrawString(aa.ToString("###,##0.000") + "(+)", fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren);//g.DrawString("-" + aa.ToString("###,##0.000"), fuente, brsh, posicioncaja("-" + aa.ToString("###,##0.000")), ren);
                        aa = 0;
                    }
                    else if (rw.Cells[1].Value.ToString().Contains("Notas de Crédito") || rw.Cells[1].Value.ToString().Contains("Otros Conceptos Comisión") || rw.Cells[1].Value.ToString().Contains("NOTA CREDITO") || rw.Cells[1].Value.ToString().Contains("Servicio de logistica USDA"))
                    {
                        g.DrawString(aa.ToString("###,##0.000") + "(-)", fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren); //g.DrawString("-" + aa.ToString("###,##0.000"), fuente, brsh, posicioncaja("-" + aa.ToString("###,##0.000")), ren);
                        aa = 0;
                    }
                    else
                        g.DrawString(aa.ToString("###,##0.000"), fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren);
                    if (z == 0)
                        totalx = totalx + 0;
                    else
                        totalx = totalx + aa;
                    z = 1;
                }

                ren = ren + 4;
            }
            g.DrawString(totalx.ToString("###,##0.000"), fuente, brsh, posicioncaja(totalx.ToString("###,##0.000")), ren);
            ren = ren + 6;
            g.DrawString("TOTAL", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.000")), ren);
            ren = ren + 4;
            g.DrawString(lbl_porcentaje.Text, fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_porcentaje.Text).ToString("###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_porcentaje.Text).ToString("###,###,##0.000")), ren);
            decimal ax = Math.Round(Convert.ToDecimal(txt_porcentaje.Text) * -1 / Convert.ToDecimal(Convert.ToDecimal(cantidad)), 2);
            g.DrawString(ax.ToString("###,###,##0.000"), fuente, brsh, 178, ren);
            ren = ren + 4;
            g.DrawString("% Desc. autoservicio:", fuente, brsh, 3, ren);
            g.DrawString("% " + txt_porce_desc.Text, fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_cant_porce.Text).ToString("###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_cant_porce.Text).ToString("###,###,##0.000")), ren);
            decimal ay = Math.Round(Convert.ToDecimal(txt_cant_porce.Text) * -1 / Convert.ToDecimal(Convert.ToDecimal(cantidad)), 2);
            g.DrawString(ay.ToString("###,###,##0.000"), fuente, brsh, 178, ren);
            ren = ren + 4;
            string tt = "";
            if (txt_tipo.Text == "NACIONAL")
                tt = "M.N.";
            else
                tt = "USD";
            g.DrawString("TOTAL A LIQUIDAR", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_liquidar.Text).ToString("###,###,###,##.000") + " " + tt, fuente, brsh, posiciontotales(Convert.ToDecimal(txt_liquidar.Text).ToString("###,###,###,##.000")), ren);
            ren = ren + 6;
            g.DrawString("COSTO UNITARIO", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_costounitario.Text).ToString("###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_costounitario.Text).ToString("###,##0.000")), ren);
            ren = ren + 6;

            g.DrawString("__________________________________________________________________________________________", fuente, brsh, 3, ren);

            ren = ren + 6;
            g.DrawString("Copia", fuente, brsh, 3, ren + 2);
            g.DrawString("Comercializadora GAB, S.A. de C.V.", fuente, brsh, 48, ren + 2);
            g.DrawString("Liquidación de Producto Terminado del " + lbl_fecha1.Text + " al " + lbl_fecha2.Text, fuente, brsh, 3, ren + 6);
            g.DrawString("Proveedor:" + lbl_cveprov.Text + " " + lbl_proveedor.Text, fuente, brsh, 3, ren + 10);
            g.DrawString("Liquidación:" + lbl_liquidacion.Text, fuente, brsh, 118, ren + 10);
            g.DrawString("Producto:" + lbl_cveprod.Text + " " + lbl_producto.Text, fuente, brsh, 3, ren + 14);
            if (txt_tipo.Text == "NACIONAL")
                g.DrawString("Nacional", fuente, brsh, 3, ren + 18);
            else
                g.DrawString("Exportación" + "Tipo cambio: " + ((txtTipoCambioResp.Text == "") ? txt_tipocambio.Text : txtTipoCambioResp.Text), fuente, brsh, 3, ren + 18);
            g.DrawString(DateTime.Now.ToLongDateString(), fuente, brsh, 118, ren + 18);
            g.DrawString("Flejes x pallet: " + lbl_flejes.Text, fuente, brsh, 48, ren + 22);
            g.DrawString("Cajas x pallet: " + lbl_cajas.Text, fuente, brsh, 118, ren + 22);
            g.DrawString("Total", fuente, brsh, 85, ren + 26);
            g.DrawString("Unidades", fuente, brsh, 82, ren + 30);
            g.DrawString("Precio", fuente, brsh, 113, ren + 26);
            g.DrawString("Unitario", fuente, brsh, 111, ren + 30);
            g.DrawString("Importe", fuente, brsh, 153, ren + 26);
            g.DrawString("Total", fuente, brsh, 154, ren + 30);
            g.DrawString("Importe", fuente, brsh, 188, ren + 26);
            g.DrawString("x caja", fuente, brsh, 189, ren + 30);

            ren = ren + 34;
            totalx = 0;
            z = 0;
            totcaj = 0;
            foreach (DataGridViewRow rw in dtgConceptos.Rows)
            {
                totcaj = Convert.ToDecimal(dtgConceptos.Rows[0].Cells[2].Value.ToString());
                totcaj2 = Convert.ToDecimal(rw.Cells[2].Value.ToString());

                if (rw.Cells[5].Value.ToString() != "")
                {
                    if (Convert.ToDecimal(rw.Cells[3].Value.ToString()) == 0)
                        continue;
                }

                //g.DrawString(rw.Cells[1].Value.ToString(), fuente, brsh, 3, ren);
                if (rw.Cells[1].Value.ToString().Trim().Length > 34)
                    g.DrawString(rw.Cells[1].Value.ToString().Substring(0, 34).ToString(), fuente, brsh, 3, ren);
                else
                    g.DrawString(rw.Cells[1].Value.ToString(), fuente, brsh, 3, ren);
                g.DrawString(rw.Cells[2].Value.ToString(), fuente, brsh, posicionunidades(rw.Cells[2].Value.ToString()), ren);
                g.DrawString(rw.Cells[3].Value.ToString(), fuente, brsh, posicionprecio(rw.Cells[3].Value.ToString()), ren);
                g.DrawString(rw.Cells[4].Value.ToString(), fuente, brsh, posicionimporte(rw.Cells[4].Value.ToString()), ren);
                if (Convert.ToDecimal(rw.Cells[4].Value.ToString()) == 0)
                    g.DrawString("0.000", fuente, brsh, posicioncaja("0.000"), ren);
                else
                {
                    decimal aa = 0;
                    if (rw.Cells[0].Value.ToString().Length > 2)
                    {
                        string val2 = verificarunidad2(rw.Cells[0].Value.ToString());
                        if (val2 == "ROL" || val2 == "ROLLO")
                        {
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / totcaj2);
                        }
                        else
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / totcaj);
                    }
                    else
                    {
                        if (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) == 0)
                            aa = 0;
                        else
                            aa = (Math.Abs(Convert.ToDecimal(rw.Cells[4].Value.ToString())) / Convert.ToDecimal(totcaj));
                    }

                    //g.DrawString(aa.ToString("###,##0.00"), fuente, brsh, posicioncaja(aa.ToString("###,##0.00")), ren);
                    if (rw.Cells[1].Value.ToString().Contains("Notas de Cargo") || rw.Cells[1].Value.ToString().Contains("NOTAS DE CARGO"))
                    {
                        g.DrawString(aa.ToString("###,##0.000") + "(+)", fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren);//g.DrawString("-" + aa.ToString("###,##0.000"), fuente, brsh, posicioncaja("-" + aa.ToString("###,##0.000")), ren);
                        aa = 0;
                    }
                    else if (rw.Cells[1].Value.ToString().Contains("Notas de Crédito") || rw.Cells[1].Value.ToString().Contains("Otros Conceptos Comisión") || rw.Cells[1].Value.ToString().Contains("NOTA CREDITO") || rw.Cells[1].Value.ToString().Contains("NOTAS DE CREDITO") || rw.Cells[1].Value.ToString().Contains("Servicio de logistica USDA"))
                    {
                        g.DrawString(aa.ToString("###,##0.000") + "(-)", fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren); //g.DrawString("-" + aa.ToString("###,##0.000"), fuente, brsh, posicioncaja("-" + aa.ToString("###,##0.000")), ren);
                        aa = 0;
                    }
                    else
                        g.DrawString(aa.ToString("###,##0.000"), fuente, brsh, posicioncaja(aa.ToString("###,##0.000")), ren);

                    if (z == 0)
                        totalx = totalx + 0;
                    else
                        totalx = totalx + aa;
                    z = 1;
                }

                ren = ren + 4;
            }
            g.DrawString(totalx.ToString("###,##0.000"), fuente, brsh, posicioncaja(totalx.ToString("###,##0.000")), ren);
            ren = ren + 6;
            g.DrawString("TOTAL", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_total.Text).ToString("###,###,###,##0.000")), ren);
            ren = ren + 4;
            g.DrawString(lbl_porcentaje.Text, fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_porcentaje.Text).ToString("###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_porcentaje.Text).ToString("###,###,##0.000")), ren);
            ax = Math.Round(Convert.ToDecimal(txt_porcentaje.Text) * -1 / Convert.ToDecimal(cantidad), 2);
            g.DrawString(ax.ToString("###,###,##0.000"), fuente, brsh, 178, ren);
            ren = ren + 4;
            g.DrawString("% Desc. autoservicio:", fuente, brsh, 3, ren);
            g.DrawString("% " + txt_porce_desc.Text, fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_cant_porce.Text).ToString("###,###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_cant_porce.Text).ToString("###,###,##0.000")), ren);
            ay = Math.Round(Convert.ToDecimal(txt_cant_porce.Text) * -1 / Convert.ToDecimal(cantidad), 2);
            g.DrawString(ay.ToString("###,###,##0.000"), fuente, brsh, 178, ren);
            ren = ren + 4;
            tt = "";
            if (txt_tipo.Text == "NACIONAL")
                tt = "M.N.";
            else
                tt = "USD";
            g.DrawString("TOTAL A LIQUIDAR", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_liquidar.Text).ToString("###,###,###,##.000") + " " + tt, fuente, brsh, posiciontotales(Convert.ToDecimal(txt_liquidar.Text).ToString("###,###,###,##.000")), ren);
            ren = ren + 6;
            g.DrawString("COSTO UNITARIO", fuente, brsh, 78, ren);
            g.DrawString(Convert.ToDecimal(txt_costounitario.Text).ToString("###,##0.000"), fuente, brsh, posiciontotales(Convert.ToDecimal(txt_costounitario.Text).ToString("###,##0.000")), ren);
        }

        private void btnImprime_Click(object sender, EventArgs e)
        {
            printDocument1.Print();
        }

        public int posicionunidades(string cad)
        {
            int pos = 0;
            if (cad.Length == 5)
                pos = 89;
            if (cad.Length == 6)
                pos = 87;
            if (cad.Length == 7)
                pos = 85;
            if (cad.Length == 9)
                pos = 81;
            if (cad.Length == 10)
                pos = 79;
            if (cad.Length == 11)
                pos = 77;
            if (cad.Length == 12)
                pos = 75;
            if (cad.Length == 13)
                pos = 73;
            if (cad.Length == 14)
                pos = 71;
            if (cad.Length == 15)
                pos = 69;
            return pos;
        }
        public int posicionprecio(string cad)
        {
            int pos = 0;
            if (cad.Length == 5)
                pos = 116;
            if (cad.Length == 6)
                pos = 114;
            if (cad.Length == 7)
                pos = 112;
            if (cad.Length == 9)
                pos = 111;
            if (cad.Length == 10)
                pos = 106;
            if (cad.Length == 11)
                pos = 104;
            if (cad.Length == 12)
                pos = 102;
            if (cad.Length == 14)
                pos = 98;
            if (cad.Length == 15)
                pos = 96;
            return pos;
        }
        public int posicionimporte(string cad)
        {
            int pos = 0;
            if (cad.Length == 5)
                pos = 156;
            if (cad.Length == 6)
                pos = 154;
            if (cad.Length == 7)
                pos = 152;
            if (cad.Length == 8)
                pos = 150;
            if (cad.Length == 9)
                pos = 149;
            if (cad.Length == 10)
                pos = 146;
            if (cad.Length == 11)
                pos = 144;
            if (cad.Length == 12)
                pos = 142;
            if (cad.Length == 13)
                pos = 140;
            if (cad.Length == 14)
                pos = 138;
            if (cad.Length == 15)
                pos = 136;
            return pos;
        }
        public int posicioncaja(string cad)
        {
            int pos = 0;
            if (cad.Length == 4)
                pos = 194;
            if (cad.Length == 5)
                pos = 192;
            if (cad.Length == 6)
                pos = 190;
            if (cad.Length == 7)
                pos = 192;
            if (cad.Length == 8)
                pos = 190;
            if (cad.Length == 9)
                pos = 191;
            if (cad.Length == 10)
                pos = 186;
            if (cad.Length == 11)
                pos = 184;
            if (cad.Length == 12)
                pos = 182;
            if (cad.Length == 13)
                pos = 180;
            if (cad.Length == 14)
                pos = 178;
            if (cad.Length == 15)
                pos = 176;
            return pos;
        }
        public int posiciontotales(string cad)
        {
            int pos = 0;
            if (cad.Length == 4)
                pos = 158;
            if (cad.Length == 5)
                pos = 156;
            if (cad.Length == 6)
                pos = 154;
            if (cad.Length == 7)
                pos = 152;
            if (cad.Length == 8)
                pos = 150;
            if (cad.Length == 9)
                pos = 148;
            if (cad.Length == 10)
                pos = 146;
            if (cad.Length == 11)
                pos = 144;
            if (cad.Length == 12)
                pos = 142;
            if (cad.Length == 13)
                pos = 140;
            if (cad.Length == 14)
                pos = 138;
            if (cad.Length == 15)
                pos = 136;
            return pos;
        }

        private void btnTexto_Click(object sender, EventArgs e)
        {
            textoliquidacion();
        }

        private void btnCancela_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea cancelar la liquidación " + txt_tipo.Text + "?", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            {
                try
                {
                    //Se hace la cancelacion
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT liq_folio, status, liq_afecto, liq_numoc1 FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                    reader1 = cmnd1.ExecuteReader();
                    bool fnd = false;

                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        if (reader1.GetValue(3).ToString().Trim() != "")
                        {
                            MessageBox.Show("La liquidación no puede ser cancelada por que ya tiene Orde de Compra,\ndebe cancelar la orden", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                            thisConnection.Close();
                            return;
                        }
                        if (reader1.GetValue(1).ToString() == "C")
                        {
                            MessageBox.Show("La liquidación esta cancelada, no puede ser cancelada nuevamente", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                            thisConnection.Close();
                            return;
                        }
                        if (reader1.GetValue(2).ToString() == "1")
                        {
                            MessageBox.Show("No puedes cancelar la liquidación, ya afecto los costos", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                            thisConnection.Close();
                            return;
                        }
                        fnd = true;
                    }
                    else
                    {
                        MessageBox.Show("La liquidación no fué encontrada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();
                        thisConnection.Close();
                        return;
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();


                    decimal naci = 0;
                    decimal expi = 0;

                    if (fnd == true)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT uni_nac, uni_exp FROM tb_mstr_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            naci = reader1.GetDecimal(0);
                            expi = reader1.GetDecimal(1);
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        if (txt_tipo.Text == "NACIONAL")
                        {
                            if (expi > 0) //SOLO SE LIMPIAN CAMPOS
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET uni_nac = '0', liq_pre_uni = '0', liq_porcen1 = '0', liq_porcen2 = '0', " +
                                    "liq_imp_tot = '0', liq_imp_por = '0', liq_imp_liq = '0', " +
                                    "liq_costo1 = '0', liq_costo2 = '0', " +
                                    "liq_por_des = '0', liq_imp_pordes = '0', status = 'A'" +
                                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE tipo_con = 'N' AND liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }
                            else //SE PROCEDE A LA CANCELACION
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET status = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_liquidacion_rec SET status = 'C' WHERE liquidacion = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                string cvep = "";
                                cvep = lbl_cveprod.Text;



                                if (txtTL.Text == "PTC")
                                {
                                    //Designa los recibos que se relacionaron para la liquidacion
                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numliq = ' ', hrp_numoc = ' ', hrp_cvepro = 'LIQCAN'" + //hrp_liquidado = ' '
                                        " WHERE hrp_numliq = '" + lbl_liquidacion.Text + "' AND lin_clave = '" + txt_lincve.Text + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                if (txtTL.Text == "PRO")
                                {
                                    if (recibs.Rows.Count > 0)
                                    {
                                        cmnd1 = thisConnection.CreateCommand();
                                        cmnd1.CommandText = "UPDATE tb_det_liq_planta SET estatus = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                        cmnd1.ExecuteNonQuery();
                                        cmnd1.Dispose();
                                    }
                                }
                            }

                            //----------14/09/2017----------//
                            //SE CONSULTA EL PRESTAMO(tb_det_prestamo) PARA SACAR LAS CANTIDADES(tb_det_prestamo[cantidad]) Y RESTARLAS A Tb_Prestamos_Prov(Saldo)
                            foreach (DataGridViewRow rA in dtgConceptos.Rows)
                            {
                                if (rA.Cells["valor"].Value.ToString() == "")
                                    continue;
                                if (Convert.ToDecimal(rA.Cells["precio"].Value.ToString()) == 0)
                                    continue;
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "SELECT cantidad FROM tb_det_prestamo WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "' " +
                                    "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N' AND estatus = 'A' AND tipo_mov = 'LQ'";
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    while (reader1.Read())
                                    {
                                        if (Convert.ToDecimal(rA.Cells["tc"].Value.ToString()) > 1)//MOVIMIENTO EN DOLARES CONVERTIR A DOLARES PARA RESTAR
                                        {
                                            string cantidade = reader1["cantidad"].ToString().Trim();
                                            string cantidade2 = (Convert.ToDecimal(cantidade) / Convert.ToDecimal(rA.Cells["tc"].Value.ToString())).ToString("0.0000");
                                            cmnd2 = thisConnection.CreateCommand();
                                            cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade2 + "' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "'";
                                            cmnd2.ExecuteNonQuery();
                                            cmnd2.Dispose();
                                        }
                                        if (Convert.ToDecimal(rA.Cells["tc"].Value.ToString()) == 1)//MOVIMIENTO EN PESOS
                                        {
                                            string cantidade = reader1["cantidad"].ToString().Trim();
                                            cmnd2 = thisConnection.CreateCommand();
                                            cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade + "' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "'";
                                            cmnd2.ExecuteNonQuery();
                                            cmnd2.Dispose();
                                        }


                                        cmnd2 = thisConnection.CreateCommand();
                                        cmnd2.CommandText = "UPDATE tb_det_prestamo SET estatus = 'C' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "' " +
                                            "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'N' AND estatus = 'A' AND tipo_mov = 'LQ'";
                                        cmnd2.ExecuteNonQuery();
                                        cmnd2.Dispose();
                                    }
                                }
                                reader1.Close();
                                reader1.Dispose();
                                cmnd1.Dispose();
                            }
                            //----------FIN 14/09/2017----------//

                            if (lblOrdenCompra.Text != "-")
                            {
                                //28/07/2021
                                //CANCELACION DEL REGISTRO DE ORDEN DE COMPRA ANTICIPADA
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_anticipada_pt SET liq_estatus = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'NACIONAL'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                //11/08/2021
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND status_oc = 'A'";
                                //cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                //16/08/2021
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET unidad_oc = '' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }

                            //desvincular liquidacion con nota de credito y cargo
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT count(nc_folio) as notas FROM tb_det_notascyc WHERE liq_folio_nal = '" + lbl_liquidacion.Text + "'";
                            reader1 = cmnd1.ExecuteReader();
                            Int32 conteo = 0;
                            if (reader1.HasRows)
                            {
                                reader1.Read();
                                conteo = Convert.ToInt32(reader1["notas"].ToString());
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();

                            if (conteo > 0)
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '0' WHERE liq_folio_nal = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }

                        }
                        else //EXPORTACION
                        {
                            if (naci > 0) //SOLO SE LIMPIAN CAMPOS
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET uni_exp = '0', liq_preunie = '0', liq_porcen_1e = '0', liq_porcen_2e = '0', " +
                                    "liq_imp_tote = '0', liq_imp_pore = '0', liq_imp_liqe = '0', " +
                                    "liq_costo1e = '0', liq_costo2e = '0', " +
                                    "liq_exp_pordes = '0', liq_exp_imppordes = '0', status = 'A', liq_precambio = '0' " +
                                    "WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "DELETE FROM tb_det_liquidacion WHERE tipo_con = 'E' AND liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }
                            else //SE PROCEDE CON LA CANCELACION
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_liquidacion SET status = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_liquidacion_rec SET status = 'C' WHERE liquidacion = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                string cvep = "";
                                cvep = lbl_cveprod.Text;

                                if (txtTL.Text == "PTC")
                                {
                                    //Designa los recibos que se relacionaron para la liquidacion
                                    cmnd1 = thisConnection.CreateCommand();
                                    cmnd1.CommandText = "UPDATE tb_hist_recepcion SET hrp_numliq = ' ', hrp_numoc = ' ', hrp_cvepro = 'LIQCAN'" + //hrp_liquidado = ' '
                                        " WHERE hrp_numliq = '" + lbl_liquidacion.Text + "' AND lin_clave = '" + txt_lincve.Text + "'";
                                    cmnd1.ExecuteNonQuery();
                                    cmnd1.Dispose();
                                }
                                if (txtTL.Text == "PRO")
                                {
                                    if (recibs.Rows.Count > 0)
                                    {
                                        cmnd1 = thisConnection.CreateCommand();
                                        cmnd1.CommandText = "UPDATE tb_det_liq_planta SET estatus = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "'";
                                        cmnd1.ExecuteNonQuery();
                                        cmnd1.Dispose();
                                    }
                                }
                            }

                            //----------14/09/2017----------//
                            //SE CONSULTA EL PRESTAMO(tb_det_prestamo) PARA SACAR LAS CANTIDADES(tb_det_prestamo[cantidad]) Y RESTARLAS A Tb_Prestamos_Prov(Saldo)
                            foreach (DataGridViewRow rA in dtgConceptos.Rows)
                            {
                                if (rA.Cells["valor"].Value.ToString() == "")
                                    continue;
                                if (Convert.ToDecimal(rA.Cells["precio"].Value.ToString()) == 0)
                                    continue;

                                string num_val = rA.Cells["valor"].Value.ToString();
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "SELECT cantidad FROM tb_det_prestamo WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "' " +
                                    "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E' AND estatus = 'A' AND tipo_mov = 'LQ'";
                                reader1 = cmnd1.ExecuteReader();
                                if (reader1.HasRows)
                                {
                                    while (reader1.Read())
                                    {
                                        if (Convert.ToDecimal(rA.Cells["tc"].Value.ToString()) > 1)//MOVIMIENTO EN DOLARES CONVERTIR A PESOS PARA RESTAR
                                        {
                                            string cantidade = reader1["cantidad"].ToString().Trim();
                                            string cantidade2 = (Convert.ToDecimal(cantidade) * Convert.ToDecimal(rA.Cells["tc"].Value.ToString())).ToString("0.0000");
                                            cmnd2 = thisConnection.CreateCommand();
                                            cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade2 + "' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "'";
                                            cmnd2.ExecuteNonQuery();
                                            cmnd2.Dispose();
                                        }
                                        if (Convert.ToDecimal(rA.Cells["tc"].Value.ToString()) == 1)//MOVIMIENTO EN DOLARES
                                        {
                                            string cantidade = reader1["cantidad"].ToString().Trim();
                                            cmnd2 = thisConnection.CreateCommand();
                                            cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade + "' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "'";
                                            cmnd2.ExecuteNonQuery();
                                            cmnd2.Dispose();
                                        }


                                        cmnd2 = thisConnection.CreateCommand();
                                        cmnd2.CommandText = "UPDATE tb_det_prestamo SET estatus = 'C' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "' " +
                                            "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E' AND estatus = 'A' AND tipo_mov = 'LQ'";
                                        cmnd2.ExecuteNonQuery();
                                        cmnd2.Dispose();
                                        //string cantidade = reader1["cantidad"].ToString().Trim();
                                        //cmnd2 = thisConnection.CreateCommand();
                                        //cmnd2.CommandText = "UPDATE Tb_Prestamos_Prov SET Saldo = Saldo - '" + cantidade + "' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "'";
                                        //cmnd2.ExecuteNonQuery();
                                        //cmnd2.Dispose();

                                        //cmnd2 = thisConnection.CreateCommand();
                                        //cmnd2.CommandText = "UPDATE tb_det_prestamo SET estatus = 'C' WHERE Id_Movimiento = '" + rA.Cells["valor"].Value.ToString() + "' " +
                                        //    "AND liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'E' AND estatus = 'A' AND tipo_mov = 'LQ'";
                                        //cmnd2.ExecuteNonQuery();
                                        //cmnd2.Dispose();
                                    }
                                }
                                reader1.Close();
                                reader1.Dispose();
                                cmnd1.Dispose();
                            }
                            //----------FIN 14/09/2017----------//


                            if (lblOrdenCompra.Text != "-")
                            {
                                //28/07/2021
                                //CANCELACION DEL REGISTRO DE ORDEN DE COMPRA ANTICIPADA
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_anticipada_pt SET liq_estatus = 'C' WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND liq_tipo = 'EXPORTACION'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                //11/08/2021
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_mstr_ordencompra SET liquidacion = '' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND status_oc = 'A'";
                                //cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();

                                //16/08/2021
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_ordenescompra SET unidad_oc = '' WHERE numero_oc = '" + lblOrdenCompra.Text + "' AND prod_clave = '" + lbl_cveprod.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }

                            //desvincular liquidacion con nota de credito y cargo
                            cmnd1 = thisConnection.CreateCommand();
                            cmnd1.CommandText = "SELECT count(nc_folio) as notas FROM tb_det_notascyc WHERE liq_folio_exp = '" + lbl_liquidacion.Text + "'";
                            reader1 = cmnd1.ExecuteReader();
                            Int32 conteo = 0;
                            if (reader1.HasRows)
                            {
                                reader1.Read();
                                conteo = Convert.ToInt32(reader1["notas"].ToString());
                            }
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();

                            if (conteo > 0)
                            {
                                cmnd1 = thisConnection.CreateCommand();
                                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '0' WHERE liq_folio_exp = '" + lbl_liquidacion.Text + "'";
                                cmnd1.ExecuteNonQuery();
                                cmnd1.Dispose();
                            }
                        }



                    }
                    thisConnection.Close();

                    string filelog = "C:\\SisEmpWeb\\eventlog.txt";
                    using (StreamWriter sw = File.AppendText(filelog))
                    {
                        sw.WriteLine(Utilerias.Class1.Usu_login.Trim() + " " + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString() + " Cancelación de liquidacion: " + lbl_liquidacion.Text);
                        sw.Close();
                    }

                    Utilerias.Class1.registrar_movimiento(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login, "B", "4.1", lbl_liquidacion.Text, "CANCELACION DE LIQUIDACION: " + lbl_liquidacion.Text, "SISEMP");

                    MessageBox.Show("Liquidación cancelada", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //BUSCAR MOVIMIENTOS DE ANTICIPO DE LIQUIDACION CANCELADA
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    cmnd1.CommandText = "SELECT Id_Movimiento FROM tb_det_liquidacion WHERE liq_folio = '" + lbl_liquidacion.Text + "' AND Id_Movimiento <> '' AND precio_con > '0'";
                    reader1 = cmnd1.ExecuteReader();
                    DataTable dtAnti = new DataTable();
                    dtAnti.Columns.Add("movi", typeof(string));
                    if (reader1.HasRows)
                    {
                        while (reader1.Read())
                        {
                            DataRow rt = dtAnti.NewRow();
                            rt["movi"] = reader1["Id_Movimiento"].ToString().Trim();
                            dtAnti.Rows.Add(rt);
                        }
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();

                    if (dtAnti.Rows.Count > 0)
                    {
                        string cad = correo_movimientos(dtAnti);
                        enviarcorreo_cancel(cad);
                    }


                    thisConnection.Close();

                    lbl_estatus.Text = "Cancelada";

                    DialogResult = System.Windows.Forms.DialogResult.OK;

                    //BORRAR LIQUIDACION DE SERVIDOR
                    FileInfo liq_copy = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
                    if (liq_copy.Exists == true)
                    {
                        liq_copy.Delete();
                    }

                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    thisConnection.Close();
                    Utilerias.Class1.registro_errores(DateTime.Now, Utilerias.Class1.Usu_login, Environment.MachineName, "5.3", ex.ToString(), "SISEMP");
                    Utilerias.Class1.SendMail("aescamilla@mrlucky.com.mx", "aescamilla", "atrejo", Environment.MachineName + " " + ex.ToString());
                }
            }
        }

        private void dtgConceptos_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            decimal cantcamb = 0;
            decimal cantcamb2 = 0;

            decimal cant_saldo = 0;
            if (e.ColumnIndex == 2)
            {
                if (e.RowIndex == 0)
                {
                    cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                    cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());
                    dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                    dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2), 3).ToString("###,###,##0.000");
                    tcon.Rows[e.RowIndex][2] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                    tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                    calculatotales();
                    calculoporcentaje();
                    if (Convert.ToDecimal(txt_porce_desc.Text) > 0)
                    {
                        KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                        txt_porce_desc_KeyPress(sender, llave);
                    }
                }
                else
                {
                    if (dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "93" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "95" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "100" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "102" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "103" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "104" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "105" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "106" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "107")
                    {
                        cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                        cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());
                        dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2), 3).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][2] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                        calculatotales();
                        calculoporcentaje();
                    }
                    else
                    {
                        cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                        cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());
                        dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2) * -1, 3).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][2] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                        calculatotales();
                        calculoporcentaje();
                    }

                }
            }
            if (e.ColumnIndex == 3)
            {
                if (e.RowIndex == 0)
                {
                    cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                    cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());
                    dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                    dtgConceptos.CurrentRow.Cells[3].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                    dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2), 3).ToString("###,###,##0.000");

                    tcon.Rows[e.RowIndex][3] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                    tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                    calculatotales();
                    calculoporcentaje();
                    if (Convert.ToDecimal(txt_porce_desc.Text) > 0)
                    {
                        KeyPressEventArgs llave = new KeyPressEventArgs(Convert.ToChar(13));
                        txt_porce_desc_KeyPress(sender, llave);
                    }
                }
                else
                {
                    if (dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "93" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "100" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "102" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "103" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "104" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "105" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "106" || dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "107")//|| dtgConceptos.CurrentRow.Cells[0].Value.ToString() == "95"
                    {
                        cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                        cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());

                        if (dtgConceptos.CurrentRow.Cells[5].Value.ToString() != "")
                        {
                            cant_saldo = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[8].Value.ToString());
                            if (cantcamb2 > cant_saldo)
                            {
                                MessageBox.Show("La cantidad ingresada es mayor al saldo restante, favor de verificar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                cantcamb2 = cant_saldo;
                                dtgConceptos.CurrentRow.Cells[3].Value = cant_saldo.ToString("###,###,##0.000");
                            }
                        }

                        dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[3].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2), 3).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][3] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                        calculatotales();
                        calculoporcentaje();
                    }
                    else
                    {
                        cantcamb = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString());
                        cantcamb2 = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString());

                        if (dtgConceptos.CurrentRow.Cells[5].Value.ToString() != "")
                        {
                            cant_saldo = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[8].Value.ToString());
                            if (cantcamb2 > cant_saldo)
                            {
                                MessageBox.Show("La cantidad ingresada es mayor al saldo restante, favor de verificar", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                cantcamb2 = cant_saldo;
                                dtgConceptos.CurrentRow.Cells[3].Value = cant_saldo.ToString("###,###,##0.000");
                            }
                        }

                        dtgConceptos.CurrentRow.Cells[2].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[2].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[3].Value = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                        dtgConceptos.CurrentRow.Cells[4].Value = Math.Round((cantcamb * cantcamb2) * -1, 3).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][3] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[3].Value.ToString()).ToString("###,###,##0.000");
                        tcon.Rows[e.RowIndex][4] = Convert.ToDecimal(dtgConceptos.CurrentRow.Cells[4].Value.ToString()).ToString("###,###,##0.000");
                        calculatotales();
                        calculoporcentaje();
                    }

                }
            }
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            printDocument1.PrinterSettings.PrinterName = "Foxit Reader PDF Printer";
            printDocument1.Print();

            FileInfo archivo = new FileInfo(@"c:\\Reportes\document.pdf");

            FileInfo liq_copy = new FileInfo(@"c:\\Reportes\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            if (liq_copy.Exists == true)
            {
                liq_copy.Delete();
            }
            archivo.CopyTo(@"c:\\Reportes\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf", true);

            FileInfo liq_copy2 = new FileInfo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
            if (liq_copy2.Exists == true)
            {
                liq_copy2.Delete();
            }
            archivo.CopyTo(@"\\\\gabira1\\liquidaciones\\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");

            Process.Start(@"c:\\Reportes\" + lbl_liquidacion.Text + "_" + txt_tipo.Text + ".pdf");
        }

        private void btnCalcula_Click(object sender, EventArgs e)
        {
            if (txt_tipocambio.Text == "")
            {
                MessageBox.Show("Debe ingresar el tipo de cambio", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            bool bnd = validarnumero(txt_tipocambio.Text);

            if (bnd == false)
            {
                MessageBox.Show("El valor introducido no es númerico", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txt_tipocambio.Text = "";
                return;
            }
            else
            {
                if (Convert.ToDecimal(txt_tipocambio.Text) == 0)
                {
                    MessageBox.Show("El valor introducido debe ser mayor a cero", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txt_tipocambio.Text = "";
                    return;
                }
            }

            if (txtTipoCambioResp.Text != "")
            {
                for (int i = 0; i < tcon.Rows.Count; i++)
                {
                    if (tcon.Rows[i]["cve_con"].ToString() == "113")
                    {
                    }

                    if (tcon.Rows[i]["valor"].ToString() != "")
                        continue;
                    if (tcon.Rows[i]["cve_con"].ToString() != "1" && tcon.Rows[i]["cve_con"].ToString() != "7" && tcon.Rows[i]["cve_con"].ToString() != "92" && tcon.Rows[i]["cve_con"].ToString() != "93" && tcon.Rows[i]["cve_con"].ToString() != "102" && tcon.Rows[i]["cve_con"].ToString() != "103" && tcon.Rows[i]["cve_con"].ToString() != "104" && tcon.Rows[i]["cve_con"].ToString() != "105" && tcon.Rows[i]["cve_con"].ToString() != "106" && tcon.Rows[i]["cve_con"].ToString() != "107" && tcon.Rows[i]["cve_con"].ToString() != "108" && tcon.Rows[i]["cve_con"].ToString() != "109")
                    {



                        if (procedencia == "EXPORTACION")
                        {
                            if (tcon.Rows[i]["cve_con"].ToString() != "6")
                            {
                                if (tcon.Rows[i]["cve_con"].ToString() == "82")
                                {
                                    if (lbl_producto.Text.Contains("BROCCOLI") || lbl_producto.Text.Contains("BROCOLI") || lbl_producto.Text.Contains("COLIFLOR"))
                                    {

                                    }
                                    else
                                    {
                                        tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) * Convert.ToDecimal(txtTipoCambioResp.Text), 3).ToString("0.000");
                                        tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                                        dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                                        dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");
                                    }
                                }
                                else
                                {
                                    if (tcon.Rows[i]["cve_con"].ToString() == "99" || tcon.Rows[i]["cve_con"].ToString() == "100" || tcon.Rows[i]["cve_con"].ToString() == "110" || tcon.Rows[i]["cve_con"].ToString() == "113")
                                    {
                                    }
                                    else
                                    {
                                        tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) * Convert.ToDecimal(txtTipoCambioResp.Text), 3).ToString("0.000");
                                        tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                                        dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                                        dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");
                                    }

                                }

                            }
                        }
                        else
                        {
                            tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) * Convert.ToDecimal(txtTipoCambioResp.Text), 3).ToString("0.000");
                            tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                            dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                            dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");
                        }
                    }
                }
            }

            for (int i = 0; i < tcon.Rows.Count; i++)
            {
                if (tcon.Rows[i]["valor"].ToString() != "")
                    continue;
                if (tcon.Rows[i]["cve_con"].ToString() == "113")
                {
                }
                if (tcon.Rows[i]["cve_con"].ToString() != "1" && tcon.Rows[i]["cve_con"].ToString() != "7" && tcon.Rows[i]["cve_con"].ToString() != "92" && tcon.Rows[i]["cve_con"].ToString() != "93" && tcon.Rows[i]["cve_con"].ToString() != "100" && tcon.Rows[i]["cve_con"].ToString() != "102" && tcon.Rows[i]["cve_con"].ToString() != "103" && tcon.Rows[i]["cve_con"].ToString() != "104" && tcon.Rows[i]["cve_con"].ToString() != "105" && tcon.Rows[i]["cve_con"].ToString() != "106" && tcon.Rows[i]["cve_con"].ToString() != "107" && tcon.Rows[i]["cve_con"].ToString() != "108" && tcon.Rows[i]["cve_con"].ToString() != "109")
                {
                    if (tcon.Rows[i]["cve_con"].ToString() == "110")
                    {
                    }
                    if (procedencia == "EXPORTACION")
                    {
                        if (tcon.Rows[i]["cve_con"].ToString() != "6")
                        {
                            if (tcon.Rows[i]["cve_con"].ToString() == "82")
                            {
                                if (lbl_producto.Text.Contains("BROCCOLI") || lbl_producto.Text.Contains("BROCOLI") || lbl_producto.Text.Contains("COLIFLOR"))
                                {

                                }
                                else
                                {

                                    tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(dtgConceptos.Rows[i].Cells["precio"].Value.ToString()), 3);

                                    tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) / Convert.ToDecimal(txt_tipocambio.Text), 3);
                                    tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                                    dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                                    dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");


                                }
                            }
                            else
                            {
                                if (tcon.Rows[i]["cve_con"].ToString() == "99" || tcon.Rows[i]["cve_con"].ToString() == "100" || tcon.Rows[i]["cve_con"].ToString() == "110" || tcon.Rows[i]["cve_con"].ToString() == "113")
                                {
                                }
                                else
                                {
                                    if (tcon.Rows[i]["cve_con"].ToString() == "111")
                                    {
                                    }
                                    tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(dtgConceptos.Rows[i].Cells["precio"].Value.ToString()), 3);

                                    tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) / Convert.ToDecimal(txt_tipocambio.Text), 3);
                                    tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                                    dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                                    dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");
                                }

                            }

                        }
                    }
                    else
                    {
                        tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(dtgConceptos.Rows[i].Cells["precio"].Value.ToString()), 3);

                        tcon.Rows[i]["precio"] = Math.Round(Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()) / Convert.ToDecimal(txt_tipocambio.Text), 3);
                        tcon.Rows[i]["total"] = Math.Round((Convert.ToDecimal(tcon.Rows[i]["unidades"].ToString()) * Convert.ToDecimal(tcon.Rows[i]["precio"].ToString())) * -1, 3);

                        dtgConceptos.Rows[i].Cells["precio"].Value = Convert.ToDecimal(tcon.Rows[i]["precio"].ToString()).ToString("0.000");
                        dtgConceptos.Rows[i].Cells["total"].Value = Convert.ToDecimal(tcon.Rows[i]["total"].ToString()).ToString("0.000");
                    }
                }
            }

            txtTipoCambioResp.Text = txt_tipocambio.Text;

            calculatotales();
        }

        private void dtgConceptos_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                if (e.RowIndex > 0)
                    return;
                if (lbl_producto.Text.Contains("TOMATE") == true)
                {
                    DetallePrecio dlg = new DetallePrecio(lbl_cveprod.Text, lbl_fecha1.Text, lbl_fecha2.Text);
                    dlg.ShowDialog();
                }

            }
            if (e.ColumnIndex == 5)
            {
                if (dtgConceptos.CurrentRow.Cells[5].Value.ToString() != "")
                {
                    string anti = "";
                    string cve_pro = "";
                    string nom_prv = "";
                    anti = dtgConceptos.CurrentRow.Cells[5].Value.ToString();
                    cve_pro = lbl_cveprov.Text;
                    nom_prv = lbl_proveedor.Text;
                    anticipos dlg = new anticipos(anti, cve_pro, nom_prv);
                    dlg.ShowDialog();
                }
            }
            if (e.ColumnIndex == 1)
            {
                if (dtgConceptos.CurrentRow.Cells[9].Value.ToString() != "1")
                    return;
                string concepto = dtgConceptos.CurrentRow.Cells[1].Value.ToString();
                string clave = dtgConceptos.CurrentRow.Cells[0].Value.ToString();
                string precio = dtgConceptos.CurrentRow.Cells[3].Value.ToString();
                string producto = lbl_producto.Text;
                string tipo = txt_tipo.Text;
                string tipo_cambio = txt_tipocambio.Text;
                string qry = "";
                if (clave.Length > 4) //Empaques
                {
                    qry = "SELECT RTRIM(hrp_recibo) AS Folio, RTRIM(emp_clave) AS Empaque, CONVERT(VARCHAR,CAST(hrp_cantidad AS MONEY),1) AS Cantidad, " +
                        "FORMAT(hrp_costo, 'C', 'es-MX') AS Costo, FORMAT((hrp_cantidad * hrp_costo), 'C', 'es-MX') AS Importe, hrp_fecha AS Fecha FROM tb_historico_recepcion " +
                        "WHERE hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR' " +
                        "AND hrp_costo > 0 AND hrp_estatus <> 'C' AND " +
                        "(hrp_fecha >= '" + lbl_fecha1.Text + "' AND hrp_fecha <= '" + lbl_fecha2.Text + "') " +
                        "AND alm_clave in ('01', '02') AND emp_clave = '" + clave + "' " +
                        "ORDER BY hrp_tipo_recepcion, emp_clave";
                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "empaque");

                    if (ds.Tables["empaque"].Rows.Count == 0)
                    {
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = "SELECT emp_clave, FORMAT(emp_costo, 'C', 'es-MX') AS Costo FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
                        reader1 = cmnd1.ExecuteReader();
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            MessageBox.Show("***** COSTO EMPAQUE *****\nNo se encontraron movimientos para calculo de costo.\nUltimo costo registrado: " + reader1["Costo"], "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            reader1.Close();
                            reader1.Dispose();
                            cmnd1.Dispose();
                            thisConnection.Close();
                            return;
                        }
                    }
                    thisConnection.Close();
                    detalle dlg = new detalle(ds, clave, producto, precio, tipo, tipo_cambio, "1");
                    dlg.ShowDialog();
                }
                if (clave == "1") //Total de cajas
                {
                    if (tipo == "NACIONAL")
                    {
                        qry = "select b.fcn_folio AS Folio, RTRIM(a.fcn_elaboro) AS Elaboro, RTRIM(b.prod_clave) AS Clave_Prod, CONVERT(VARCHAR,CAST(b.fcn_num_unidades AS MONEY),1) AS Unidades, " +
                        "FORMAT(b.fcn_precio_mn, 'C', 'es-MX') AS Precio_MN, FORMAT((b.fcn_precio_mn * b.fcn_num_unidades), 'C', 'es-MX') AS Importe, a.fcn_fecha AS Fecha " +
                        "from tb_mstr_facturas_nal a, tb_det_facturas b " +
                        "where (a.fcn_fecha >= '" + lbl_fecha1.Text + "' AND a.fcn_fecha <= '" + lbl_fecha2.Text + "') " +
                        "and a.fcn_estatus <> 'C' and  b.fcn_folio = a.fcn_folio AND a.fcn_monto <> a.ncr_monto " +
                        "and b.fcn_tipo = a.fcn_lugar AND a.um_clave = 'PESOS' AND b.prod_clave = '" + lbl_cveprod.Text + "' " +
                        "order by b.prod_clave, a.fcn_fecha";
                    }
                    else
                    {
                        qry = "select b.fcn_folio AS Folio, RTRIM(a.fcn_elaboro) AS Elaboro, RTRIM(b.prod_clave) AS Clave_Prod, CONVERT(VARCHAR,CAST(b.fcn_num_unidades AS MONEY),1) AS Unidades, " +
                        "FORMAT(b.fcn_precio_usd, 'C', 'es-MX') AS Precio_MN, FORMAT((b.fcn_precio_usd * b.fcn_num_unidades), 'C', 'es-MX') AS Importe, a.fcn_fecha AS Fecha " +
                        "from tb_mstr_facturas_nal a, tb_det_facturas b " +
                        "where (a.fcn_fecha >= '" + lbl_fecha1.Text + "' AND a.fcn_fecha <= '" + lbl_fecha2.Text + "') " +
                        "and a.fcn_estatus <> 'C' and  b.fcn_folio = a.fcn_folio AND a.fcn_monto <> a.ncr_monto " +
                        "and b.fcn_tipo = a.fcn_lugar AND a.um_clave = 'USD' AND b.prod_clave = '" + lbl_cveprod.Text + "' " +
                        "order by b.prod_clave, a.fcn_fecha";
                    }

                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "total_de_cajas");
                    thisConnection.Close();

                    if (ds.Tables["total_de_cajas"].Rows.Count == 0)
                    {
                        MessageBox.Show("No se detectaron ventas para ese producto", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "2")//Tarimas
                {
                    qry = "select h.hrp_recibo AS Entrada, RTRIM(h.emp_clave) AS Clave, RTRIM(e.emp_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(h.hrp_cantidad AS MONEY),1) AS Unidades, " +
                        "FORMAT(h.hrp_costo, 'C', 'es-MX') AS Costo, FORMAT((h.hrp_cantidad * h.hrp_costo), 'C', 'es-MX') AS Importe " +
                        "from tb_historico_recepcion h, tb_cat_empaques e " +
                        "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C' " +
                        "and h.hrp_fecha >= '" + lbl_fecha1.Text + "' and h.hrp_fecha <= '" + lbl_fecha2.Text + "' AND e.emp_nombre like 'TARIMA%'";
                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "tarimas");
                    thisConnection.Close();
                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "3")//Enfriamiento No Aplica
                { }
                if (clave == "4")//Flejes No Aplica
                { }
                if (clave == "5")//Esquineros
                {
                    qry = "select h.hrp_recibo AS Entrada, RTRIM(h.emp_clave) AS Clave, RTRIM(e.emp_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(h.hrp_cantidad AS MONEY),1) AS Unidades, " +
                        "FORMAT(h.hrp_costo, 'C', 'es-MX') AS Costo, FORMAT((h.hrp_cantidad * h.hrp_costo), 'C', 'es-MX') AS Importe " +
                        "from tb_historico_recepcion h, tb_cat_empaques e " +
                        "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' and hrp_estatus <> 'C' " +
                        "and h.hrp_fecha >= '" + lbl_fecha1.Text + "' and h.hrp_fecha <= '" + lbl_fecha2.Text + "' AND e.emp_nombre like 'ESQUINERO%'";
                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "esquineros");
                    thisConnection.Close();
                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "6")//Fletes
                {
                    string mes_det = "";
                    string anio_det = "";
                    DateTimeFormatInfo fe1 = new CultureInfo("es-ES", false).DateTimeFormat;
                    if (Convert.ToDateTime(lbl_fecha1.Text) > Convert.ToDateTime(lbl_fecha2.Text))
                        mes_det = fe1.GetMonthName(Convert.ToDateTime(lbl_fecha2.Text).Month);
                    else
                        mes_det = fe1.GetMonthName(Convert.ToDateTime(lbl_fecha1.Text).Month);
                    anio_det = Convert.ToDateTime(lbl_fecha1.Text).Year.ToString();
                    thisConnection.Open();
                    cmnd1 = thisConnection.CreateCommand();
                    if (tipo == "NACIONAL")
                    {
                        cmnd1.CommandText = "SELECT mes, FORMAT(costo, 'C', 'es-MX') AS Costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'NAL' " +
                        "AND mes = '" + mes_det + "' AND año = '" + anio_det + "' ORDER BY prod_clave";
                    }
                    else
                    {
                        cmnd1.CommandText = "SELECT mes, FORMAT(costo, 'C', 'es-MX') AS Costo FROM tb_cat_costosprod2 WHERE prod_clave = '" + lbl_cveprod.Text + "' AND movimiento = 'EXP' " +
                        "AND mes = '" + mes_det + "' AND año = '" + anio_det + "' ORDER BY prod_clave";
                    }
                    reader1 = cmnd1.ExecuteReader();
                    if (reader1.HasRows)
                    {
                        reader1.Read();
                        MessageBox.Show("***** COSTO POR FLETE *****\nProducto: " + producto + "\nMes: " + reader1["mes"].ToString() + "\nCosto: " + reader1["Costo"].ToString(), "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    reader1.Close();
                    reader1.Dispose();
                    cmnd1.Dispose();
                    thisConnection.Close();
                }
                if (clave == "7")//Mermas y Reclamaciones
                {
                    if (tipo == "NACIONAL")
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep AS Clave, RTRIM(A.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, " +
                            "FORMAT(A.dnc_precio_mn, 'C', 'es-MX') AS Costo, FORMAT((A.dnc_cantidad * A.dnc_precio_mn), 'C', 'es-MX') AS Importe, fechap AS Fecha " +
                            "FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + lbl_cveprov.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.clavep = '" + lbl_cveprod.Text + "' and A.nc_folio = B.nc_folio AND A.dnc_lugar = B.nc_lugar and B.nc_estatus <> 'C' ORDER BY A.nc_Folio";
                    }
                    else
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep AS Clave, RTRIM(A.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, " +
                            "FORMAT(A.dnc_precio_usd, 'C', 'es-MX') AS Costo, FORMAT((A.dnc_cantidad * A.dnc_precio_usd), 'C', 'es-MX') AS Importe, fechap AS Fecha " +
                            "FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + lbl_cveprov.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.clavep = '" + lbl_cveprod.Text + "' and A.nc_folio = B.nc_folio AND A.dnc_lugar = B.nc_lugar and B.nc_estatus <> 'C' ORDER BY A.nc_Folio";
                    }

                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "mermas");
                    thisConnection.Close();

                    if (ds.Tables["mermas"].Rows.Count == 0)
                    {
                        MessageBox.Show("No se encontraron movimiento para calculo de Mermas y Reclamaciones.", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "92")//Notas de Crédito x Dif. en Precio
                {
                    if (tipo == "NACIONAL")
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep, A.prod_nombre, " +
                            "CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, FORMAT(A.dnc_precio_usd, 'C', 'es-MX') AS Costo, " +
                            "FORMAT((A.dnc_cantidad * A.dnc_precio_mn), 'C', 'es-MX') AS Importe, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.cveprov = '" + lbl_cveprov.Text + "' AND A.clavep = '" + lbl_cveprod.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.lin_clave = '9803' AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_lugar = B.nc_lugar " +
                            "AND A.dnc_tipo = B.nc_tipo AND A.fechap = B.nc_fecha";
                    }
                    else
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep, A.prod_nombre, " +
                            "CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, FORMAT(A.dnc_precio_usd, 'C', 'es-MX') AS Costo, " +
                            "FORMAT((A.dnc_cantidad * A.dnc_precio_usd), 'C', 'es-MX') AS Importe, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.clavep = '" + lbl_cveprod.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.lin_clave = '9803' AND A.dnc_tipo = 'NCR' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_lugar = B.nc_lugar " +
                            "AND A.dnc_tipo = B.nc_tipo AND A.fechap = B.nc_fecha";
                    }
                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "NCDF");
                    thisConnection.Close();
                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "8")//Rechazos por Calidad No Aplica
                { }
                if (clave == "83")//FUMIGACIONES
                {
                    if (tipo == "EXPORTACION")
                    {
                        thisConnection.Open();
                        qry = "select RTRIM(h.emp_clave) AS Clave, RTRIM(e.emp_nombre) AS Descripcion , CONVERT(VARCHAR,CAST(SUM(h.hrp_cantidad) AS MONEY),1) AS Unidades, FORMAT(SUM(h.hrp_cantidad * h.hrp_costo), 'C', 'es-mx') as Importe " +
                            "from tb_historico_recepcion h, tb_cat_empaques e " +
                            "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
                            "and h.hrp_fecha >= '" + lbl_fecha1.Text + "' and h.hrp_fecha <= '" + lbl_fecha2.Text + "' AND h.emp_clave in ('M2628', 'N3742') " +
                            "GROUP BY h.emp_clave, e.emp_nombre";
                        SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                        DataSet ds = new DataSet();
                        adapter.Fill(ds, "Historico");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + lbl_fecha1.Text + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%APIO%' AND DF.lin_clave in ('01', '16') AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "Apio");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + lbl_fecha1.Text + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%KALE%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto AND DF.prod_clave NOT IN ('09009ESK28', '16KAML1220', '16KAOML121', '16KAORCH42') " +
                            "GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "Kale");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + lbl_fecha1.Text + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%LECHUGA%OREJONA%' AND DF.lin_clave IN ('05', '09', '16') AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND DF.prod_clave NOT IN ('09HOLEOR33', '09HOORML25', '09HOLEOR25', '16001HLO12', '09LEROCH62', '09009LOH14', '09009LEO41', " +
                            "'09009LEO45', '09009LOB62', '05005LO1X4', '05005LETOR', '05005LETAY', '09TALEOJ41') " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "Orejona");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + lbl_fecha1.Text + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%SWISS%CHARD%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "Swiss");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + lbl_fecha1.Text + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%RAINBOW%CHARD%' AND DF.lin_clave = '16' AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "Rainbow");

                        qry = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp ORDER BY invemp_fecha DESC";
                        cmnd1 = thisConnection.CreateCommand();
                        cmnd1.CommandText = qry;
                        reader1 = cmnd1.ExecuteReader();
                        string fecha_cierre = "";
                        if (reader1.HasRows)
                        {
                            reader1.Read();
                            fecha_cierre = Convert.ToDateTime(reader1["invemp_fecha"]).ToShortDateString();
                        }
                        reader1.Close();
                        reader1.Dispose();
                        cmnd1.Dispose();

                        qry = "select RTRIM(h.emp_clave) AS Clave, RTRIM(e.emp_nombre) AS Descripcion , CONVERT(VARCHAR,CAST(SUM(h.hrp_cantidad) AS MONEY),1) AS Unidades " +
                            "from tb_historico_recepcion h, tb_cat_empaques e " +
                            "WHERE h.emp_clave = e.emp_clave AND h.hrp_tipo_recepcion = 'ENT' and h.hrp_situacion = 'NOR' AND h.hrp_estatus <> 'C' " +
                            "and h.hrp_fecha >= '" + fecha_cierre + "' and h.hrp_fecha <= '" + lbl_fecha2.Text + "' AND h.emp_clave in ('M2628', 'N3742') " +
                            "GROUP BY h.emp_clave, e.emp_nombre";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "HistoricoNulo");
                        qry = "SELECT RTRIM(DF.prod_clave) AS Clave, RTRIM(P.prod_nombre) AS Descripcion, CONVERT(VARCHAR,CAST(SUM(DF.fcn_num_unidades) AS MONEY),1) AS Unidades " +
                            "FROM tb_det_facturas DF,tb_mstr_facturas_nal F, tb_cat_producto P " +
                            "WHERE DF.fcn_folio = F.fcn_folio AND F.fcn_estatus <> 'C' AND F.fcn_fecha >= '" + fecha_cierre + "' and F.fcn_fecha <= '" + lbl_fecha2.Text + "' " +
                            "AND F.fcn_lugar = 'EXP'  AND DF.prod_clave = P.prod_clave AND P.prod_nombre like '%APIO%' AND DF.lin_clave in ('01', '16') AND DF.fcn_tipo = F.fcn_lugar " +
                            "AND F.fcn_monto <> F.ncr_monto GROUP BY P.prod_nombre, DF.prod_clave ORDER BY P.prod_nombre, DF.prod_clave";
                        adapter = new SqlDataAdapter(qry, thisConnection);
                        adapter.Fill(ds, "ApioNulo");

                        thisConnection.Close();
                        detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                        dlg.ShowDialog();

                    }
                }
                if (clave == "93")//Notas de Cargo
                {
                    if (tipo == "NACIONAL")
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep, A.prod_nombre, " +
                            "CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, FORMAT(A.dnc_precio_mn, 'C', 'es-MX') AS Costo, " +
                            "FORMAT((A.dnc_cantidad * A.dnc_precio_mn), 'C', 'es-MX') AS Importe, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.cveprov = '" + lbl_cveprov.Text + "' AND A.clavep = '" + lbl_cveprod.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.lin_clave = '9803' AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_lugar = B.nc_lugar" +
                            "ORDER BY A.fechap";
                    }
                    else
                    {
                        qry = "SELECT A.nc_folio AS Folio, A.clavep, A.prod_nombre, " +
                            "CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, FORMAT(A.dnc_precio_usd, 'C', 'es-MX') AS Costo, " +
                            "FORMAT((A.dnc_cantidad * A.dnc_precio_usd), 'C', 'es-MX') AS Importe, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.clavep = '" + lbl_cveprod.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.lin_clave = '9803' AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_lugar = B.nc_lugar " +
                            "ORDER BY A.fechap";
                    }


                    thisConnection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds, "NotasCargo");
                    thisConnection.Close();
                    detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                    dlg.ShowDialog();
                }
                if (clave == "100")//Servicio de logistica
                {
                    if (tipo == "EXPORTACION")
                    {
                        qry = "SELECT A.nc_folio AS Folio, RTRIM(A.clavep) AS Clave, RTRIM(A.prod_nombre) AS Descripcion, " +
                            "CONVERT(VARCHAR,CAST(A.dnc_cantidad AS MONEY),1) AS Unidades, FORMAT(A.dnc_precio_usd, 'C', 'es-MX') AS Costo, " +
                            "FORMAT((A.dnc_cantidad * A.dnc_precio_usd), 'C', 'es-MX') AS Importe, A.fechap FROM tb_det_notascyc A, tb_mstr_notascyc B " +
                            "WHERE A.clavep = '" + lbl_cveprod.Text + "' AND (A.fechap >= '" + lbl_fecha1.Text + "' AND A.fechap <= '" + lbl_fecha2.Text + "') " +
                            "AND A.lin_clave = '9812' AND A.dnc_tipo = 'NCG' and A.nc_folio = B.nc_folio and B.nc_estatus <> 'C' AND A.dnc_lugar = B.nc_lugar";
                        thisConnection.Open();
                        SqlDataAdapter adapter = new SqlDataAdapter(qry, thisConnection);
                        DataSet ds = new DataSet();
                        adapter.Fill(ds, "ServicioLogistica");
                        thisConnection.Close();
                        detalle dlg = new detalle(ds, concepto, producto, precio, tipo, tipo_cambio, "0");
                        dlg.ShowDialog();
                    }
                }
            }

        }

        public string correo_movimientos(DataTable dtA)
        {
            //DataTable dtMovs = new DataTable();
            //dtMovs.Columns.Add("Movimiento", typeof(string));
            //dtMovs.Columns.Add("Cultivo", typeof(string));
            //dtMovs.Columns.Add("Anticipo", typeof(string));
            //dtMovs.Columns.Add("Contrato", typeof(string));
            //dtMovs.Columns.Add("Fol", typeof(string));
            //dtMovs.Columns.Add("NE", typeof(string));
            //dtMovs.Columns.Add("Tip", typeof(string));
            //dtMovs.Columns.Add("Cantidad", typeof(string));
            //dtMovs.Columns.Add("Descuento", typeof(string));
            //dtMovs.Columns.Add("Saldo", typeof(string));
            if (thisConnection.State == ConnectionState.Closed)
                thisConnection.Open();
            int i = 0;
            bool deja_renglon = false;
            string cadena = "";
            foreach (DataRow t in dtA.Rows)
            {
                if (i == 1 && deja_renglon == true)
                {
                    cadena += "</ br></ br></ br></ br>";
                }
                cadena += "<p><table border='1'>";
                cadena += "<thead>";
                cadena += "<th>MOVIMIENTO</th>";
                cadena += "<th>CULTIVO</th>";
                cadena += "<th>ANTICIPO</th>";
                cadena += "<th>CONTRATO</th>";
                cadena += "<th>FOLIO</th>";
                cadena += "<th>NAC / EXP</th>";
                cadena += "<th>TIPO</th>";
                cadena += "<th>CANTIDAD</th>";
                cadena += "<th>DESCUENTO</th>";
                cadena += "<th>SALDO</th>";
                cadena += "</thead>";

                cadena += "</tbody>";
                decimal original = 0;
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT A.Fecha, A.Id_Contrato, A.Descripcion_Art, A.Cantidad, A.contrato, A.factura, B.prod_nombre, A.moneda FROM Tb_Prestamos_Prov A JOIN " +
                    "tb_cat_producto B ON B.prod_clave = A.Id_Contrato WHERE Id_Movimiento = '" + t["movi"].ToString() + "'";
                reader1 = cmnd1.ExecuteReader();
                DataRow u;
                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        //u = dtMovs.NewRow();
                        //u["Movimiento"] = t["movi"].ToString();
                        //u["Cultivo"] = reader1["prod_nombre"].ToString().Trim();
                        //u["Anticipo"] = reader1["Descripcion_Art"].ToString().Trim();
                        //u["Contrato"] = reader1["contrato"].ToString().Trim();
                        //u["Cantidad"] = reader1["Cantidad"].ToString().Trim();
                        //dtMovs.Rows.Add(u);
                        original = Convert.ToDecimal(reader1["Cantidad"].ToString().Trim());

                        cadena += "<tr>";
                        cadena += "<td>" + t["movi"].ToString() + "</td>";
                        cadena += "<td>" + reader1["prod_nombre"].ToString().Trim() + "</td>";
                        cadena += "<td>" + reader1["Descripcion_Art"].ToString().Trim() + "</td>";
                        cadena += "<td>" + reader1["contrato"].ToString().Trim() + "</td>";
                        cadena += "<td></td>";
                        cadena += "<td></td>";
                        cadena += "<td>" + reader1["moneda"].ToString().Trim() + "</td>";
                        cadena += "<td align='right'>" + Convert.ToDecimal(reader1["Cantidad"].ToString().Trim()).ToString("$###,###,##0.0000") + "</td>";
                        cadena += "<td></td>";
                        cadena += "<td></td>";
                        cadena += "</tr>";
                    }
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();



                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "SELECT liq_folio, cantidad, liq_tipo, tipo_mov, tipo_cambio FROM tb_det_prestamo WHERE Id_Movimiento = '" + t["movi"].ToString() + "' AND estatus = 'A'";
                reader1 = cmnd1.ExecuteReader();
                decimal cantidad = 0;
                DataRow rt;
                decimal sumatoria = 0;

                if (reader1.HasRows)
                {
                    while (reader1.Read())
                    {
                        //rt = dtMovs.NewRow();
                        //rt["Fol"] = reader1["liq_folio"].ToString().Trim();
                        //rt["NE"] = reader1["liq_tipo"].ToString().Trim();
                        //rt["Tip"] = reader1["tipo_mov"].ToString().Trim();
                        //rt["Cantidad"] = reader1["cantidad"].ToString().Trim();
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

                        //rt["Descuento"] = cantidad.ToString();
                        //dtMovs.Rows.Add(rt);

                        string liq_tipo = (reader1["liq_tipo"].ToString().Trim() == "N") ? "NACIONAL" : (reader1["liq_tipo"].ToString().Trim() == "N") ? "EXPORTACION" : "";
                        string tipo_mov = (reader1["tipo_mov"].ToString().Trim() == "LQ") ? "LIQUIDACION" : (reader1["tipo_mov"].ToString().Trim() == "MP") ? "OC MAT PRIMA" : (reader1["tipo_mov"].ToString().Trim() == "ES") ? "LIQ. ESPARRAGO" : "TOTAL";

                        cadena += "<tr>";
                        cadena += "<td></td>";
                        cadena += "<td></td>";
                        cadena += "<td></td>";
                        cadena += "<td></td>";
                        cadena += "<td>" + reader1["liq_folio"].ToString().Trim() + "</td>";
                        cadena += "<td>" + liq_tipo + "</td>";
                        cadena += "<td>" + tipo_mov + "</td>";
                        cadena += "<td></td>";
                        cadena += "<td align='right'>" + cantidad.ToString("$###,###,##0.0000") + "</td>";
                        cadena += "<td></td>";
                        cadena += "</tr>";

                        cantidad = 0;


                    }

                    cadena += "<tr>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td></td>";
                    cadena += "<td align='right'>" + sumatoria.ToString("$###,###,##0.0000") + "</td>";
                    cadena += "<td align='right'>" + (original - sumatoria).ToString("$###,###,##0.0000") + "</td>";
                    cadena += "</tr>";

                    //rt = dtMovs.NewRow();
                    //rt["Fol"] = "";
                    //rt["NE"] = "";
                    //rt["Tip"] = "";
                    //rt["Descuento"] = sumatoria.ToString();
                    //rt["Saldo"] = (original - sumatoria).ToString();
                    //dtMovs.Rows.Add(rt);
                }
                reader1.Close();
                reader1.Dispose();
                cmnd1.Dispose();
                cadena += "</tbody></table></p>";
                i++;
                deja_renglon = true;
            }
            //thisConnection.Close();

            return cadena;
        }

        public void enviarcorreo(string cuerpo)
        {

            string host = Dns.GetHostName();
            IPHostEntry ipEntry = Dns.GetHostEntry(host);

            System.Net.Mail.MailMessage mnsg = new System.Net.Mail.MailMessage();
            //mnsg.To.Add("cmoreno@mrlucky.com.mx");//msamano@mrlucky.com.mx
            mnsg.To.Add("aescamilla@mrlucky.com.mx");
            //mnsg.CC.Add("jcardenas@mrlucky.com.mx");
            mnsg.Subject = "Liquidación: " + lbl_liquidacion.Text + ". Movimientos de anticipos afectados";
            mnsg.SubjectEncoding = System.Text.Encoding.UTF8;
            mnsg.Body = cuerpo;
            mnsg.BodyEncoding = System.Text.Encoding.UTF8;
            mnsg.IsBodyHtml = true;
            mnsg.From = new MailAddress("aescamilla@mrlucky.com.mx");

            SmtpClient cliente = new SmtpClient();
            cliente.Credentials = new System.Net.NetworkCredential("aescamilla", "atrejo");
            cliente.Port = 587;
            cliente.EnableSsl = true;
            cliente.Host = "mail1.mrlucky.com.mx";

            try
            {
                cliente.Send(mnsg);
                //Response.Write("<script>alert('Correo enviado correctamente');</script>");
            }
            catch (Exception)
            {
                //Response.Write("<script>alert('No fue enviado el correo electronico');</script>");
            }
        }

        public void enviarcorreo_cancel(string cuerpo)
        {

            string host = Dns.GetHostName();
            IPHostEntry ipEntry = Dns.GetHostEntry(host);

            System.Net.Mail.MailMessage mnsg = new System.Net.Mail.MailMessage();
            //mnsg.To.Add("cmoreno@mrlucky.com.mx");//msamano@mrlucky.com.mx
            mnsg.To.Add("aescamilla@mrlucky.com.mx");
            //mnsg.CC.Add("jcardenas@mrlucky.com.mx");
            mnsg.Subject = "Cancelación de Liquidación: " + lbl_liquidacion.Text + ". Movimientos de anticipos afectados";
            mnsg.SubjectEncoding = System.Text.Encoding.UTF8;
            mnsg.Body = cuerpo;
            mnsg.BodyEncoding = System.Text.Encoding.UTF8;
            mnsg.IsBodyHtml = true;
            mnsg.From = new MailAddress("aescamilla@mrlucky.com.mx");

            SmtpClient cliente = new SmtpClient();
            cliente.Credentials = new System.Net.NetworkCredential("aescamilla", "atrejo");
            cliente.Port = 587;
            cliente.EnableSsl = true;
            cliente.Host = "mail1.mrlucky.com.mx";

            try
            {
                cliente.Send(mnsg);
                //Response.Write("<script>alert('Correo enviado correctamente');</script>");
            }
            catch (Exception)
            {
                //Response.Write("<script>alert('No fue enviado el correo electronico');</script>");
            }
        }

        private void preliminar_Load(object sender, EventArgs e)
        {
            this.Activated += AfterLoading;
        }

        private void AfterLoading(object sender, EventArgs e)
        {
            this.Activated -= AfterLoading;

            //MessageBox.Show("works!!");
            if (txtTipoLiq.Text == "consulta")
                return;

            thisConnection.Open();
            cmnd1 = thisConnection.CreateCommand();

            cmnd1.CommandText = "SELECT A.numero_oc, FORMAT(A.fecha_oc, 'dd-MM-yyyy') AS fecha_oc, RTRIM(B.prod_clave) AS cveprod_oc, B.nomprod_oc AS nomprod_oc, " +
                "FORMAT(B.cantidad_oc, 'N2', 'es-mx') AS cantidad_oc, " +
                "FORMAT(B.precio_oc, 'N2', 'es-mx') AS precio_oc, FORMAT(B.importe_oc, 'N2', 'es-mx') AS importe_oc, B.conse " +
                "FROM tb_mstr_ordencompra A " +
                "JOIN tb_det_ordenescompra B ON A.numero_oc = B.numero_oc " +
                "WHERE A.anticipada = '1' AND A.status_oc = 'A' AND " +
                "B.prod_clave = '" + lbl_cveprod.Text + "' AND A.cveprov_oc = '" + lbl_cveprov.Text + "' AND A.tipo_oc = '" + ((txt_tipo.Text == "NACIONAL") ? "N" : "E") + "'" +
                "ORDER BY A.fecha_oc DESC";
            reader1 = cmnd1.ExecuteReader();
            bool fnd = false;
            if (reader1.HasRows)
            {
                fnd = true;
            }
            reader1.Close();
            reader1.Dispose();
            cmnd1.Dispose();
            thisConnection.Close();

            if (fnd == true)
            {
                articipadas dlg = new articipadas(lbl_cveprov.Text, lbl_proveedor.Text, lbl_cveprod.Text, lbl_producto.Text, cantidad, txt_tipo.Text, lbl_fecha1.Text, lbl_fecha2.Text);
                dlg.ShowDialog();
                if (dlg.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    lblOrdenCompra.Text = articipadas.SharedDatos.DatosCell.ordencompra;
                    lblConse.Text = articipadas.SharedDatos.DatosCell.conse;
                    if (articipadas.SharedDatos.DatosCell.recalculo == "1")
                    {
                        chkRecalculo.Checked = true;
                    }
                }
            }

        }

        public void afecta_notas_credito_exportacion(string liq, string f1, string f2, string pr, string prv)
        {
            DataTable dt_cyc = new DataTable();
            DataTable dt_merm = new DataTable();
            DataTable dt_usda = new DataTable();
            thisConnection.Open();
            //datos de notas de credito y cargo
            SqlDataAdapter adap = new SqlDataAdapter("SELECT A.nc_folio, A.lin_clave, A.dnc_lugar, A.dnc_tipo, A.clavep, A.prod_nombre " +
                "FROM tb_det_notascyc A " +
                "INNER JOIN tb_mstr_notascyc B ON A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
                "WHERE  (A.fechap >= '" + f1 + "' AND A.fechap <= '" + f2 + "') " +
                "AND A.lin_clave in ('9803', '9813', '9814', '9815', '9812') " +
                "AND A.dnc_tipo in ('NCR', 'NCG') and B.nc_estatus <> 'C' AND A.clavep = '" + pr + "' AND A.dnc_precio_usd > 0 AND liq_folio_exp = '0' " +
                "ORDER BY cveprov, prod_clave, fechap", thisConnection);
            adap.Fill(dt_cyc);

            //datos merma
            adap = new SqlDataAdapter("SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap " +
                "FROM tb_det_notascyc A INNER JOIN tb_mstr_notascyc B ON A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.nc_folio = B.nc_folio " +
                "WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + prv + "' AND A.fechap BETWEEN '" + f1 + "' AND '" + f2 + "' AND A.clavep = '" + pr + "' AND A.dnc_precio_usd > 0 " +
                "and B.nc_estatus <> 'C' AND liq_folio_exp = '0'", thisConnection);
            adap.Fill(dt_merm);

            //daTos logistica usda
            adap = new SqlDataAdapter("SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap " +
                "FROM tb_det_notascyc A " +
                "INNER JOIN tb_mstr_notascyc B ON A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
                "WHERE  (A.fechap >= '" + f1 + "' AND A.fechap <= '" + f2 + "') " +
                "AND A.lin_clave = '9812' " +
                "AND A.prod_clave = '981218' AND A.dnc_tipo = 'NCR' and B.nc_estatus <> 'C' AND A.clavep = '" + pr + "' AND A.dnc_precio_usd > 0 AND liq_folio_exp = '0' " +
                "ORDER BY cveprov, prod_clave, fechap", thisConnection);
            adap.Fill(dt_usda);

            //ACTUALIZACION DE NCR - NOTAS DE CREDITO POR DIFERENCIA EN PRECIO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9803' AND dnc_tipo = 'NCR'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9803' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCG - NOTAS DE CARGO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9803' AND dnc_tipo = 'NCG'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCG' AND clavep = '" + pr + "' AND lin_clave = '9803' AND dnc_lugar = 'EXE'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR - NOTAS DE CREDITO X ACOND EMP DESTINO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9813' AND dnc_tipo = 'NCR'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9813' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCG - NOTAS DE CARGO X ACOND EMP ORIGEN
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9814' AND dnc_tipo = 'NCG'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCG' AND clavep = '" + pr + "' AND lin_clave = '9814' AND dnc_lugar = 'EXE'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR - OTROS CONCEPTOS COMISION
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9815' AND dnc_tipo = 'NCR'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9815' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCG - SERVICIO DE LOGISTICA
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9812' AND dnc_tipo = 'NCG'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCG' AND clavep = '" + pr + "' AND lin_clave = '9812' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR DE MERMAS
            foreach (DataRow row in dt_merm.Rows)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9801' AND prod_nombre LIKE 'MERMA%' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR DE SERVICIO DE LOGISTICA USDA
            foreach (DataRow row in dt_usda.Rows)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_exp = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9812' AND prod_clave = '981218' AND dnc_lugar = 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }
            thisConnection.Close();
        }

        public void afecta_notas_credito_nacional(string liq, string f1, string f2, string pr, string prv)
        {
            DataTable dt_cyc = new DataTable();
            DataTable dt_merm = new DataTable();
            thisConnection.Open();
            //datos de notas de credito y cargo
            SqlDataAdapter adap = new SqlDataAdapter("SELECT A.nc_folio, A.lin_clave, A.dnc_lugar, A.dnc_tipo, A.clavep, A.prod_nombre " +
                "FROM tb_det_notascyc A " +
                "INNER JOIN tb_mstr_notascyc B ON A.nc_folio = B.nc_folio AND A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar " +
                "WHERE  (A.fechap >= '" + f1 + "' AND A.fechap <= '" + f2 + "') " +
                "AND A.lin_clave in ('9803', '9813', '9814') " +
                "AND A.dnc_tipo in ('NCR', 'NCG') and B.nc_estatus <> 'C' AND A.clavep = '" + pr + "' AND A.dnc_precio_mn > 0 AND liq_folio_nal = '0' " +
                "ORDER BY cveprov, prod_clave, fechap", thisConnection);
            adap.Fill(dt_cyc);

            //datos merma
            adap = new SqlDataAdapter("SELECT A.prod_nombre, A.nc_folio, A.dnc_cantidad, A.dnc_precio_mn, A.dnc_precio_usd, A.clavep, A.dnc_tipo, A.lin_clave, A.fechap " +
                "FROM tb_det_notascyc A INNER JOIN tb_mstr_notascyc B ON A.dnc_tipo = B.nc_tipo AND A.dnc_lugar = B.nc_lugar AND A.dnc_devbon = B.nc_devbon AND A.nc_folio = B.nc_folio " +
                "WHERE A.prod_nombre LIKE 'MERMA%' AND A.cveprov = '" + prv + "' AND A.fechap BETWEEN '" + f1 + "' AND '" + f2 + "' AND A.clavep = '" + pr + "' AND A.dnc_precio_mn > 0 " +
                "and B.nc_estatus <> 'C' AND A.liq_folio_nal = '0' AND A.cveprov = '" + prv + "'", thisConnection);
            adap.Fill(dt_merm);

            //ACTUALIZACION DE NCR - NOTAS DE CREDITO POR DIFERENCIA EN PRECIO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9803' AND dnc_tipo = 'NCR'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9803' AND cveprov = '" + prv + "' AND dnc_lugar <> 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCG - NOTAS DE CARGO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9803' AND dnc_tipo = 'NCG'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCG' AND clavep = '" + pr + "' AND lin_clave = '9803' AND cveprov = '" + prv + "' AND dnc_lugar <> 'EXE'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR - NOTAS DE CREDITO X ACOND EMP DESTINO
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9813' AND dnc_tipo = 'NCR'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND lin_clave = '9813' AND cveprov = '" + prv + "' AND dnc_lugar <> 'EXP'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCG - NOTAS DE CARGO X ACOND EMP ORIGEN
            foreach (DataRow row in dt_cyc.Select("lin_clave = '9814' AND dnc_tipo = 'NCG'"))
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCG' AND clavep = '" + pr + "' AND lin_clave = '9814' AND cveprov = '" + prv + "' AND dnc_lugar <> 'EXE'";
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            //ACTUALIZACION DE NCR DE MERMAS
            foreach (DataRow row in dt_merm.Rows)
            {
                cmnd1 = thisConnection.CreateCommand();
                cmnd1.CommandText = "UPDATE tb_det_notascyc SET liq_folio_nal = '" + lbl_liquidacion.Text + "' WHERE nc_folio = '" + row["nc_folio"].ToString().Trim() + "' " +
                    "AND fechap >= '" + f1 + "' AND fechap <= '" + f2 + "' AND dnc_tipo = 'NCR' AND clavep = '" + pr + "' AND prod_nombre LIKE 'MERMA%' AND dnc_lugar <> 'EXP'";//AND lin_clave = '9801'
                cmnd1.ExecuteNonQuery();
                cmnd1.Dispose();
            }

            thisConnection.Close();
        }

        private void btnNotas_Click(object sender, EventArgs e)
        {
            afecta_notas_credito_exportacion(lbl_liquidacion.Text, lbl_fecha1.Text, lbl_fecha2.Text, lbl_cveprod.Text, lbl_cveprov.Text);
        }

        public string recalculo_costos_empaque(string clave)
        {
            string mcod = "";
            decimal minvi = 0;
            decimal mcosp = 0;

            string mrec = "";
            string mtip = "";
            string malm = "";
            decimal mcant = 0;

            decimal A = 0, B = 0, C = 0, D = 0, C2 = 0;

            DataTable tbhist = new DataTable();
            tbhist.Columns.Add("hrp_recibo", typeof(string));
            tbhist.Columns.Add("hrp_fecha", typeof(string));
            tbhist.Columns.Add("emp_clave", typeof(string));
            tbhist.Columns.Add("hrp_cantidad", typeof(string));
            tbhist.Columns.Add("hrp_tipo_recepcion", typeof(string));
            tbhist.Columns.Add("alm_clave", typeof(string));
            tbhist.Columns.Add("hrp_situacion", typeof(string));
            tbhist.Columns.Add("hrp_costo", typeof(string));
            tbhist.Columns.Add("hrp_regimen", typeof(string));
            tbhist.Columns.Add("hrp_procedencia", typeof(string));
            tbhist.Columns.Add("hrp_costo_prom", typeof(string));
            tbhist.Columns.Add("hrp_tipo", typeof(string));

            DataTable tbinv = new DataTable();
            tbinv.Columns.Add("emp_clave", typeof(string));
            tbinv.Columns.Add("invemp_inicial", typeof(string));
            tbinv.Columns.Add("invemp_salidas", typeof(string));
            tbinv.Columns.Add("invemp_entradas", typeof(string));
            tbinv.Columns.Add("alm_clave", typeof(string));
            tbinv.Columns.Add("invemp_costopro", typeof(string));

            //cat_empaques ya esa cargado en dtempaques

            DataTable tbent = new DataTable();
            tbent.Columns.Add("ent_folio", typeof(string));
            tbent.Columns.Add("emp_clave", typeof(string));
            tbent.Columns.Add("alm_clave", typeof(string));
            tbent.Columns.Add("entd_cantidad", typeof(string));
            tbent.Columns.Add("entd_costo", typeof(string));
            tbent.Columns.Add("ent_tipo", typeof(string));

            DataTable tbsal = new DataTable();
            tbsal.Columns.Add("sal_folio", typeof(string));
            tbsal.Columns.Add("emp_clave", typeof(string));
            tbsal.Columns.Add("sald_cantidad", typeof(string));
            tbsal.Columns.Add("sald_costo", typeof(string));
            tbsal.Columns.Add("alm_clave", typeof(string));
            tbsal.Columns.Add("sal_tipo", typeof(string));

            DataTable tbinv2 = new DataTable();
            tbinv2.Columns.Add("emp_clave", typeof(string));
            tbinv2.Columns.Add("invemp_inicial", typeof(string));
            tbinv2.Columns.Add("invemp_salidas", typeof(string));
            tbinv2.Columns.Add("invemp_entradas", typeof(string));
            tbinv2.Columns.Add("alm_clave", typeof(string));
            tbinv2.Columns.Add("invemp_costopro", typeof(string));

            //DateTime fecha = new DateTime();
            string fech1 = "";//Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
            //fech1 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
            string fecha_inicio = "";
            string fecha_fin = "";

            //fecha_inicio = fecha.AddMonths(-1).ToShortDateString();
            fecha_fin = Convert.ToDateTime(f2).ToShortDateString();//FECHA FIN DE RANGO DE LIQUIDACIONES   //fecha.ToShortDateString();//fecha.AddDays(-1).ToShortDateString();

            SqlDataReader reader_cto;
            SqlCommand cmnd_cto;
            SqlDataAdapter adap_cto;

            string query = "";
            string glo_var_chr_alm_clave = "";
            try
            {
                //thisConnection.Open();

                //SACAR FECHA ULTIMO CIERRE
                string var_date_fecha = "";
                cmnd_cto = thisConnection.CreateCommand();
                query = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
                cmnd_cto.CommandText = query;
                reader_cto = cmnd_cto.ExecuteReader();
                if (reader_cto.HasRows)
                {
                    reader_cto.Read();
                    var_date_fecha = reader_cto.GetValue(0).ToString().Trim();
                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();

                //var_date_fecha = "01/10/2023";

                cmnd_cto = thisConnection.CreateCommand();
                query = "SELECT alm_clave FROM tb_cat_empaques WHERE emp_clave = '" + clave + "'";
                cmnd_cto.CommandText = query;
                reader_cto = cmnd_cto.ExecuteReader();
                if (reader_cto.HasRows)
                {
                    reader_cto.Read();
                    glo_var_chr_alm_clave = reader_cto.GetValue(0).ToString().Trim();

                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();

                fecha_inicio = var_date_fecha;//FECHA INICIAL VIENE DE ULTIMO CIERRE DEL INVENTARIO

                DataTable dtSPHistoricoCierre = new DataTable();
                SqlDataAdapter adap1 = new SqlDataAdapter("spSISEMPHistoricoCierre", thisConnection);
                adap1.SelectCommand.CommandType = CommandType.StoredProcedure;
                adap1.SelectCommand.Parameters.AddWithValue("@fecha1", Convert.ToDateTime(fecha_inicio).ToShortDateString());
                adap1.SelectCommand.Parameters.AddWithValue("@fecha2", fecha_fin);
                adap1.SelectCommand.Parameters.AddWithValue("@almacen", glo_var_chr_alm_clave);
                adap1.SelectCommand.Parameters.AddWithValue("@clave", clave);
                adap1.Fill(dtSPHistoricoCierre);

                DataRow rw;
                foreach (DataRow rt in dtSPHistoricoCierre.Rows)
                {
                    rw = tbhist.NewRow();
                    rw["hrp_recibo"] = rt["hrp_recibo"].ToString().Trim();
                    rw["hrp_fecha"] = rt["hrp_fecha"].ToString().Trim();
                    rw["emp_clave"] = rt["emp_clave"].ToString().Trim();
                    rw["hrp_cantidad"] = rt["hrp_cantidad"].ToString().Trim();
                    rw["hrp_tipo_recepcion"] = rt["hrp_tipo_recepcion"].ToString().Trim();
                    rw["alm_clave"] = rt["alm_clave"].ToString().Trim();
                    rw["hrp_situacion"] = rt["hrp_situacion"].ToString().Trim();
                    rw["hrp_costo"] = rt["hrp_costo"].ToString().Trim();
                    rw["hrp_regimen"] = rt["hrp_regimen"].ToString().Trim();
                    rw["hrp_procedencia"] = rt["hrp_procedencia"].ToString().Trim();
                    rw["hrp_costo_prom"] = rt["hrp_costo_prom"].ToString().Trim();
                    rw["hrp_tipo"] = rt["hrp_tipo"].ToString().Trim();
                    tbhist.Rows.Add(rw);
                }

                query = "SELECT emp_clave, alm_clave FROM tb_cat_empaques WHERE emp_nombre <> '' AND alm_clave = '" + glo_var_chr_alm_clave + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//
                cmnd_cto.CommandText = query;
                DataRow r1;
                reader_cto = cmnd_cto.ExecuteReader();
                if (reader_cto.HasRows)
                {
                    while (reader_cto.Read())
                    {
                        r1 = tbinv.NewRow();
                        r1["emp_clave"] = reader_cto.GetValue(0).ToString().Trim();
                        r1["invemp_inicial"] = "0";
                        r1["invemp_salidas"] = "0";
                        r1["invemp_entradas"] = "0";
                        r1["alm_clave"] = reader_cto.GetValue(1).ToString().Trim();
                        r1["invemp_costopro"] = "0";
                        tbinv.Rows.Add(r1);
                    }
                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();


                query = "SELECT emp_clave, invemp_inicial, invemp_salidas, invemp_entradas, alm_clave, invemp_costopro FROM tb_mstr_inventario_emp " +
                     "WHERE invemp_fecha = '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//--and emp_clave = 'B0001'
                cmnd_cto.CommandText = query;
                reader_cto = cmnd_cto.ExecuteReader();
                DataRow r2;
                if (reader_cto.HasRows)
                {
                    while (reader_cto.Read())
                    {
                        r2 = tbinv2.NewRow();
                        r2["emp_clave"] = reader_cto.GetValue(0).ToString().Trim();//grid
                        r2["invemp_inicial"] = reader_cto.GetValue(1).ToString().Trim();//grid
                        r2["invemp_salidas"] = reader_cto.GetValue(2).ToString().Trim();//no
                        r2["invemp_entradas"] = reader_cto.GetValue(3).ToString().Trim();//no
                        r2["alm_clave"] = reader_cto.GetValue(4).ToString().Trim();//ya lo traigo
                        r2["invemp_costopro"] = reader_cto.GetValue(5).ToString().Trim();//grid
                        tbinv2.Rows.Add(r2);
                    }
                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();
                //thisConnection.Close();

                foreach (DataRow rz in tbinv.Rows)
                {
                    foreach (DataRow ry in tbinv2.Select("emp_clave = '" + rz["emp_clave"] + "'"))
                    {
                        rz["invemp_inicial"] = ry["invemp_inicial"];
                        rz["invemp_costopro"] = ry["invemp_costopro"];
                    }
                }

                //thisConnection.Open();
                bool hay = false;
                //int y = 0;
                foreach (DataRow rinv in tbinv.Rows)
                {
                    mcod = rinv[0].ToString();
                    minvi = Convert.ToDecimal(rinv[1].ToString());
                    mcosp = Convert.ToDecimal(rinv[5].ToString());



                    foreach (DataRow rhis in tbhist.Select("emp_clave = '" + mcod + "'"))
                    {
                        hay = true;
                        //mesh = Convert.ToDateTime(rhis[1].ToString()).Month.ToString();
                        mrec = rhis["hrp_recibo"].ToString();
                        mtip = rhis["hrp_tipo_recepcion"].ToString();
                        malm = rhis["alm_clave"].ToString();
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "NOR")
                        {
                            decimal hrp_cantidad = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            decimal hrp_costo = Convert.ToDecimal(rhis["hrp_costo"].ToString());
                            A = Math.Abs(minvi * mcosp);
                            B = hrp_cantidad * hrp_costo;
                            C = minvi + hrp_cantidad;
                            C2 = Math.Abs(minvi) + hrp_cantidad;
                            D = (C2 == 0) ? hrp_costo : ((A + B) / C2);
                            minvi = C;
                            mcosp = Math.Abs(D);
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "ENT" && rhis["hrp_situacion"].ToString() == "DEV")
                        {
                            minvi = minvi + Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            mcant = Convert.ToDecimal(rhis["hrp_cantidad"].ToString());


                            //cmnd1 = new SqlCommand("spSISEMPHistoricoUpdateDev", thisConnection);
                            //cmnd1.CommandType = CommandType.StoredProcedure;
                            //cmnd1.Parameters.AddWithValue("@minvi", minvi);
                            //cmnd1.Parameters.AddWithValue("@mcosp", mcosp);
                            //cmnd1.Parameters.AddWithValue("@mtip", mtip);
                            //cmnd1.Parameters.AddWithValue("@malm", malm);
                            //cmnd1.Parameters.AddWithValue("@mrec", mrec);
                            //cmnd1.Parameters.AddWithValue("@mcod", mcod);
                            //cmnd1.Parameters.AddWithValue("@fecha", Convert.ToDateTime(fecha_inicio).ToShortDateString());
                            //cmnd1.Parameters.AddWithValue("@situ", "DEV");
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();




                            decimal cto = 0;
                            cto = mcosp * mcant;

                            //cmnd1 = new SqlCommand("spSISEMPDetalleUpdateDev", thisConnection);
                            //cmnd1.CommandType = CommandType.StoredProcedure;
                            //cmnd1.Parameters.AddWithValue("@cto", cto.ToString());
                            //cmnd1.Parameters.AddWithValue("@malm", malm);
                            //cmnd1.Parameters.AddWithValue("@mrec", mrec);
                            //cmnd1.Parameters.AddWithValue("@mcod", mcod);
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();
                        }
                        if (rhis["hrp_tipo_recepcion"].ToString() == "SAL")
                        {
                            minvi = minvi - Convert.ToDecimal(rhis["hrp_cantidad"].ToString());
                            //Actualizacion de historico
                            //cmnd1 = new SqlCommand("spSISEMPHistoricoUpdateDev", thisConnection);
                            //cmnd1.CommandType = CommandType.StoredProcedure;
                            //cmnd1.Parameters.AddWithValue("@minvi", minvi);
                            //cmnd1.Parameters.AddWithValue("@mcosp", mcosp);
                            //cmnd1.Parameters.AddWithValue("@mtip", mtip);
                            //cmnd1.Parameters.AddWithValue("@malm", malm);
                            //cmnd1.Parameters.AddWithValue("@mrec", mrec);
                            //cmnd1.Parameters.AddWithValue("@mcod", mcod);
                            //cmnd1.Parameters.AddWithValue("@fecha", Convert.ToDateTime(fecha_inicio).ToShortDateString());
                            //cmnd1.Parameters.AddWithValue("@situ", "NOR");
                            //cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                            //cmnd1 = new SqlCommand("spSISEMPDetalleUpdateSal", thisConnection);
                            //cmnd1.CommandType = CommandType.StoredProcedure;
                            //cmnd1.Parameters.AddWithValue("@mcosp", mcosp.ToString());
                            //cmnd1.Parameters.AddWithValue("@malm", malm);
                            //cmnd1.Parameters.AddWithValue("@mrec", mrec);
                            //cmnd1.Parameters.AddWithValue("@mcod", mcod);
                            ////cmnd1.ExecuteNonQuery();
                            //cmnd1.Dispose();

                        }
                    }

                }
                //thisConnection.Close();

                string filtro = "hrp_tipo_recepcion = 'ENT' AND hrp_situacion = 'NOR'";
                DataRow[] filasNormal = tbhist.Select(filtro);
                int contador = filasNormal.Length;
                if (contador == 0)
                    mcosp = 0;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //if (thisConnection.State == ConnectionState.Open)
                //thisConnection.Close();
            }

            return mcosp.ToString();
        }

        public string recalculo_costos_empaque_inicial(string clave)
        {
            string mcod = "";
            decimal minvi = 0;
            decimal mcosp = 0;

            string mrec = "";
            string mtip = "";
            string malm = "";
            decimal mcant = 0;

            decimal A = 0, B = 0, C = 0, D = 0, C2 = 0;

            //DateTime fecha = new DateTime();
            string fech1 = "";//Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
            //fech1 = Convert.ToDateTime(dtpFecha2.Text).ToShortDateString();
            string fecha_inicio = "";
            string fecha_fin = "";

            //fecha_inicio = fecha.AddMonths(-1).ToShortDateString();
            fecha_fin = Convert.ToDateTime(f2).ToShortDateString();//FECHA FIN DE RANGO DE LIQUIDACIONES   //fecha.ToShortDateString();//fecha.AddDays(-1).ToShortDateString();

            SqlDataReader reader_cto;
            SqlCommand cmnd_cto;
            SqlDataAdapter adap_cto;

            string query = "";
            string glo_var_chr_alm_clave = "";
            try
            {
                //thisConnection.Open();

                //SACAR FECHA ULTIMO CIERRE
                string var_date_fecha = "";
                cmnd_cto = thisConnection.CreateCommand();
                query = "SELECT TOP 1 invemp_fecha FROM tb_mstr_inventario_emp WHERE invemp_fecha <= '" + DateTime.Now.ToShortDateString() + "' ORDER BY invemp_fecha DESC";
                cmnd_cto.CommandText = query;
                reader_cto = cmnd_cto.ExecuteReader();
                if (reader_cto.HasRows)
                {
                    reader_cto.Read();
                    var_date_fecha = reader_cto.GetValue(0).ToString().Trim();
                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();

                fecha_inicio = var_date_fecha;//FECHA INICIAL VIENE DE ULTIMO CIERRE DEL INVENTARIO

                query = "SELECT emp_clave, invemp_inicial, invemp_salidas, invemp_entradas, alm_clave, invemp_costopro FROM tb_mstr_inventario_emp " +
                     "WHERE invemp_fecha = '" + Convert.ToDateTime(fecha_inicio).ToShortDateString() + "' and emp_clave = '" + clave + "' ORDER BY emp_clave";//--and emp_clave = 'B0001'
                cmnd_cto.CommandText = query;
                reader_cto = cmnd_cto.ExecuteReader();
                DataRow r2;
                if (reader_cto.HasRows)
                {
                    while (reader_cto.Read())
                    {
                        mcosp = reader_cto.GetDecimal(5);
                    }
                }
                reader_cto.Close();
                reader_cto.Dispose();
                cmnd_cto.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error de sistema", "SISEMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return mcosp.ToString();
        }

        public decimal costo_promedio_caja_coliflor(string emp1, string emp2, string fch1, string fch2)
        {
            DataTable dtCP = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter("spSISEMPLiquidacionesCostoPromedioEmpaque", thisConnection);
            adp.SelectCommand.CommandType = CommandType.StoredProcedure;
            adp.SelectCommand.Parameters.AddWithValue("@emp1", emp1);
            adp.SelectCommand.Parameters.AddWithValue("@emp2", emp2);
            adp.SelectCommand.Parameters.AddWithValue("@fch1", fch1);
            adp.SelectCommand.Parameters.AddWithValue("@fch2", fch2);
            adp.Fill(dtCP);
            return Math.Round(Convert.ToDecimal(dtCP.Rows[0]["CostoPromedio"]), 3);
        }


    }
}
