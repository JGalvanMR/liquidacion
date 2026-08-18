using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using liquidacion.Controllers;

namespace liquidacion
{
    public partial class Notas_Credito_Cargo : Form
    {
        private ConsultaController _consultaController;
        DataTable dtFiltradoLineas = new DataTable();
        DataTable dtFiltradoLineasN = new DataTable();
        public Notas_Credito_Cargo(DataTable dt, string f1, string f2, string pr, DataTable dtLineas)
        {
            InitializeComponent();

            lblClave.Text = pr;
            lblFecha1.Text = f1;
            lblFecha2.Text = f2;



            _consultaController = new ConsultaController();

            DataTable dtNCRNCGExp = _consultaController.ConsultaNCRNCGExp(f1, f2, pr);
            dtFiltradoLineas = dtNCRNCGExp.Clone();
            foreach (DataRow rL in dtLineas.Rows)
            {
                var rows = dtNCRNCGExp.AsEnumerable().Where(r => r.Field<string>("linea").Contains(rL["lin_cve"].ToString()) && r.Field<Int32>("Liquidacion") == 0);
                if (rows.Any())
                {
                    foreach (var row in rows)
                    {
                        dtFiltradoLineas.ImportRow(row); // Importamos cada fila encontrada
                    }
                }
            }

            DataTable dtNCRNCGNal = _consultaController.ConsultaNCRNCGNal(f1, f2, pr);
            dtFiltradoLineasN = dtNCRNCGNal.Clone();
            foreach (DataRow rL in dtLineas.Rows)
            {
                var rows = dtNCRNCGNal.AsEnumerable().Where(r => r.Field<string>("Linea").Contains(rL["lin_cve"].ToString()) && r.Field<Int32>("Liquidacion") == 0);
                if (rows.Any())
                {
                    foreach (var row in rows)
                    {
                        dtFiltradoLineasN.ImportRow(row); // Importamos cada fila encontrada
                    }
                }
            }

            dtgNacional.DataSource = dtFiltradoLineasN;

            dtgExportacion.DataSource = dtFiltradoLineas;


            DataTable dtProductos = _consultaController.ConsultaProductos(f1, f2, pr);
            dtgProductos.DataSource = dtProductos;



            #region anterior

            ////DataTable dtNacional = _consultaController.ConsultaNotasNacionales(f1, f2, "", pr);
            //DataTable dtExportacion = _consultaController.ConsultaNotasExportacion(f1, f2, "");

            //DataTable dtExportacionMerma = _consultaController.ConsultaNotasExportacionMerma(f1, f2, "", pr);
            ////DataTable dtNacionalMerma = _consultaController.ConsultaNotasNacionalMerma(f1, f2, "", pr);

            ////DataTable dtNacionalNC = dtNacional.Copy();
            ////dtNacionalNC.Clear();

            //DataTable dtExportacionNC = dtExportacion.Copy();
            //dtExportacionNC.Clear();

            //foreach (DataRow rp in dtProductos.Rows)
            //{
            //    if (dtExportacion.Rows.Count > 0)
            //    {
            //        foreach (DataRow re in dtExportacion.Select("Clave_Prod = '" + rp["Cve_Prod"].ToString() + "'"))
            //        {
            //            dtExportacionNC.Rows.Add(re.ItemArray);
            //        }
            //    }
            //    if (dtExportacionMerma.Rows.Count > 0)
            //    {
            //        foreach (DataRow re in dtExportacionMerma.Select("Clave_Prod = '" + rp["Cve_Prod"].ToString() + "'"))
            //        {
            //            dtExportacionNC.Rows.Add(re.ItemArray);
            //        }
            //    }


            //    //if (dtNacional.Rows.Count > 0)
            //    //{
            //    //    foreach (DataRow re in dtNacional.Select("Clave_Prod = '" + rp["Cve_Prod"].ToString() + "'"))
            //    //    {
            //    //        dtNacionalNC.Rows.Add(re.ItemArray);
            //    //    }
            //    //}
            //    //if (dtNacionalMerma.Rows.Count > 0)
            //    //{
            //    //    foreach (DataRow re in dtNacionalMerma.Select("Clave_Prod = '" + rp["Cve_Prod"].ToString() + "'"))
            //    //    {
            //    //        dtNacionalNC.Rows.Add(re.ItemArray);
            //    //    }
            //    //}
            //}



            ////DataTable dtFilterNal = new DataTable();
            ////DataTable dtFilterExp = new DataTable();

            ////dtFilterNal = dtNacional.Copy();
            ////dtFilterNal.Clear();

            ////dtFilterExp = dtExportacion.Copy();
            ////dtFilterExp.Clear();

            ////string prod_act = "";
            ////string prod_ant = "";
            ////foreach (DataRow rw in dtNacional.Rows)
            ////{
            ////    prod_act = rw["Clave_Prod"].ToString();
            ////    if (prod_ant != prod_act)
            ////    {
            ////        bool fnd = false;
            ////        foreach (DataRow rt in dt.Select("pro_clave = '" + prod_act + "'"))
            ////        {
            ////            fnd = true;
            ////        }
            ////        if (fnd == false)
            ////        {
            ////            foreach (DataRow rw2 in dtNacional.Select("Clave_Prod = '" + prod_act + "'"))
            ////            {
            ////                dtFilterNal.Rows.Add(rw2.ItemArray);
            ////            }
            ////        }
            ////    }
            ////    prod_ant = prod_act;
            ////}

            ////prod_act = "";
            ////prod_ant = "";
            ////foreach (DataRow rw in dtExportacion.Rows)
            ////{
            ////    prod_act = rw["Clave_Prod"].ToString();
            ////    if (prod_ant != prod_act)
            ////    {
            ////        bool fnd = false;
            ////        foreach (DataRow rt in dt.Select("pro_clave = '" + prod_act + "'"))
            ////        {
            ////            fnd = true;
            ////        }
            ////        if (fnd == false)
            ////        {
            ////            foreach (DataRow rw2 in dtExportacion.Select("Clave_Prod = '" + prod_act + "'"))
            ////            {
            ////                dtFilterExp.Rows.Add(rw2.ItemArray);
            ////            }
            ////        }
            ////    }
            ////    prod_ant = prod_act;
            ////}

            //dtgNacional.DataSource = dtNacionalNC;
            //dtgExportacion.DataSource = dtExportacionNC;

            #endregion

        }

        private void Notas_Credito_Cargo_Load(object sender, EventArgs e)
        {

        }
    }
}
