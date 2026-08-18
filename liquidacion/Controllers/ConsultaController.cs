using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using liquidacion.Models;
using System.Data;

namespace liquidacion.Controllers
{
    public class ConsultaController
    {
        private ConsultaModels _consultaModel;

        public ConsultaController()
        {
            _consultaModel = new ConsultaModels();
        }

        public DataTable ConsultaNotasNacionales(string f1, string f2, string pr, string pv)
        {
            return _consultaModel.ConsultaNotasNacionales(f1, f2, pr, pv);
        }

        public DataTable ConsultaNotasExportacion(string f1, string f2, string pr)
        {
            return _consultaModel.ConsultaNotasExportacion(f1, f2, pr);
        }

        public DataTable ConsultaNotasExportacionMerma(string f1, string f2, string pr, string pv)
        {
            return _consultaModel.ConsultaNotasExportacionMerma(f1, f2, pr, pv);
        }

        public DataTable ConsultaNotasNacionalMerma(string f1, string f2, string pr, string pv)
        {
            return _consultaModel.ConsultaNotasNacionalMerma(f1, f2, pr, pv);
        }

        public DataTable ConsultaProductos(string f1, string f2, string pr)
        {
            return _consultaModel.ConsultaProductos(f1, f2, pr);
        }

        public DataTable ConsultaNCRNCGExp(string f1, string f2, string pr)
        {
            return _consultaModel.ConsultaNCRNCGExp(f1, f2, pr);
        }

        public DataTable ConsultaNCRNCGNal(string f1, string f2, string pr)
        {
            return _consultaModel.ConsultaNCRNCGNal(f1, f2, pr);
        }
    }
}
