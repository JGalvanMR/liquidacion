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
    public partial class detalle : Form
    {
        public detalle(DataSet ds, string concepto, string producto, string precio, string tipo, string tipo_cambio, string empaque)
        {
            InitializeComponent();

            lblConcepto.Text = concepto;
            lblProducto.Text = producto;
            lblPrecio.Text = precio;
            lblTipo.Text = tipo;
            lblTipoCambio.Text = tipo_cambio;

            DataTable dtFumigaciones = new DataTable();
            dtFumigaciones.Columns.Add("Clave", typeof(string));
            dtFumigaciones.Columns.Add("Descripcion", typeof(string));
            dtFumigaciones.Columns.Add("Unidades", typeof(string));
            dtFumigaciones.Columns.Add("Importe", typeof(string));

            if (empaque != "1")
            {
                if (concepto == "Total de Cajas")
                {
                    decimal suma_unidades = ds.Tables["total_de_cajas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades").Replace(",", "")));
                    decimal suma_importe = ds.Tables["total_de_cajas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["total_de_cajas"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["total_de_cajas"];
                    dtgDetalle.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Tarimas")
                {
                    decimal suma_unidades = ds.Tables["tarimas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades").Replace(",", "")));
                    decimal suma_importe = ds.Tables["tarimas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["tarimas"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["tarimas"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Enfriamiento")
                {
                }
                if (concepto == "Flejes")
                {
                }
                if (concepto == "Esquineros")
                {
                    decimal suma_unidades = ds.Tables["esquineros"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades")));
                    decimal suma_importe = ds.Tables["esquineros"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["esquineros"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["esquineros"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Fletes")//Se envia mensaje en pantalla preliminar
                {
                }
                if (concepto == "Mermas y Reclamaciones")
                {
                    decimal suma_unidades = ds.Tables["mermas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades")));
                    decimal suma_importe = ds.Tables["mermas"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["mermas"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["mermas"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Notas de Crédito x Dif. en Precio")
                {
                    decimal suma_unidades = ds.Tables["NCDF"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades")));
                    decimal suma_importe = ds.Tables["NCDF"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["NCDF"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["NCDF"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Rechazos por Calidad")
                {
                }
                if (concepto == "Notas de Cargo")
                {
                    decimal suma_unidades = ds.Tables["NotasCargo"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades")));
                    decimal suma_importe = ds.Tables["NotasCargo"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["NotasCargo"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["NotasCargo"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "FUMIGACIONES")
                {
                    DataRow rw;
                    Decimal suma_hist = 0;
                    Decimal suma_fact = 0;
                    foreach (DataRow rh in ds.Tables["Historico"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString(), rh["Importe"].ToString());
                        suma_hist = Convert.ToDecimal(rh["Importe"].ToString().Replace("$", ""));
                    }
                    foreach (DataRow rh in ds.Tables["Apio"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                        suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                    }
                    foreach (DataRow rh in ds.Tables["Kale"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                        suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                    }
                    foreach (DataRow rh in ds.Tables["Orejona"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                        suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                    }
                    foreach (DataRow rh in ds.Tables["Swiss"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                        suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                    }
                    foreach (DataRow rh in ds.Tables["Rainbow"].Rows)
                    {
                        dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                        suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                    }

                    if (suma_hist == 0 || suma_fact == 0)
                    {
                        foreach (DataRow rh in ds.Tables["HistoricoNulo"].Rows)
                        {
                            dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString(), rh["Importe"].ToString());
                            suma_hist = Convert.ToDecimal(rh["Importe"].ToString().Replace("$", ""));
                        }
                        foreach (DataRow rh in ds.Tables["ApioNulo"].Rows)
                        {
                            dtFumigaciones.Rows.Add(rh["Clave"].ToString(), rh["Descripcion"].ToString(), rh["Unidades"].ToString());
                            suma_fact = suma_fact + Convert.ToDecimal(rh["Unidades"].ToString());
                        }
                    }

                    decimal precio_fum = 0;
                    precio_fum = Math.Round((suma_hist / suma_fact), 3);
                    dtFumigaciones.Rows.Add("", "Totales", suma_fact.ToString("###,##0.000"), suma_hist.ToString("###,##0.000"));
                    dtFumigaciones.Rows.Add("", "Precio Fumigacion", precio_fum.ToString("$##0.000"));
                    dtgDetalle.DataSource = dtFumigaciones;
                    dtgDetalle.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                if (concepto == "Servicio de logistica")
                {
                    decimal suma_unidades = ds.Tables["ServicioLogistica"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Unidades")));
                    decimal suma_importe = ds.Tables["ServicioLogistica"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                    decimal precio_venta = suma_importe / suma_unidades;
                    ds.Tables["ServicioLogistica"].Rows.Add(null, null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.00"), suma_importe.ToString("$###,###,##0.00"));
                    dtgDetalle.DataSource = ds.Tables["ServicioLogistica"];
                    dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dtgDetalle.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            else
            {
                decimal suma_unidades = ds.Tables["empaque"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Cantidad")));
                decimal suma_importe = ds.Tables["empaque"].AsEnumerable().Sum(r => Convert.ToDecimal(r.Field<string>("Importe").Replace("$", "")));
                decimal precio_venta = suma_importe / suma_unidades;
                ds.Tables["empaque"].Rows.Add(null, "TOTALES", suma_unidades.ToString("###,##0"), precio_venta.ToString("$###,##0.000"), suma_importe.ToString("$###,###,##0.00"), null);
                dtgDetalle.DataSource = ds.Tables["empaque"];
                dtgDetalle.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dtgDetalle.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dtgDetalle.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dtgDetalle.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

        }

        private void detalle_Load(object sender, EventArgs e)
        {

        }
    }
}
