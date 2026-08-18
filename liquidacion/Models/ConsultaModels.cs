using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace liquidacion.Models
{
    public class ConsultaModels
    {
        public DataTable ConsultaNotasNacionales(string f1, string f2, string pr, string pv)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNotasCreditoCargoNal", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@f1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@f2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@pr", pr);
                adapter.SelectCommand.Parameters.AddWithValue("@pv", pv);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }

        public DataTable ConsultaNotasNacionalMerma(string f1, string f2, string pr, string pv)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNotasCreditoCargoMermaNal", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@f1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@f2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@pr", pr);
                adapter.SelectCommand.Parameters.AddWithValue("@pv", pv);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }


        //TRAE TODAS LAS NOTAS DE CREDITO DE EXPORTACION DEL PRODUCTO
        public DataTable ConsultaNotasExportacion(string f1, string f2, string pr)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNotasCreditoCargoExp", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@f1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@f2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@pr", pr);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }

        public DataTable ConsultaNotasExportacionMerma(string f1, string f2, string pr, string pv)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNotasCreditoCargoMermaExp", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@f1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@f2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@pr", pr);
                adapter.SelectCommand.Parameters.AddWithValue("@pv", pv);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }

        //TRAE TODOS LOS PRODUCTOS A LIQUIDAR
        public DataTable ConsultaProductos(string f1, string f2, string pr)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesProductos", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@f1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@f2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@pr", pr);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }

        public DataTable ConsultaNCRNCGExp(string f1, string f2, string pr)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNCRNCGExp", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@fc1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@fc2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@prv", pr);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }

        public DataTable ConsultaNCRNCGNal(string f1, string f2, string pr)
        {
            DataTable dtConsulta = new DataTable();
            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                SqlDataAdapter adapter = new SqlDataAdapter("spSISEMPLiquidacionesNCRNCGNal", conn);
                adapter.SelectCommand.CommandType = CommandType.StoredProcedure;
                adapter.SelectCommand.Parameters.AddWithValue("@fc1", f1);
                adapter.SelectCommand.Parameters.AddWithValue("@fc2", f2);
                adapter.SelectCommand.Parameters.AddWithValue("@prv", pr);
                adapter.Fill(dtConsulta);
            }
            return dtConsulta;
        }
    }
}
