using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ScrapClarov1.Funciones;

namespace ScrapClarov1
{
    public partial class V1 : Form
    {
        public V1()
        {
            InitializeComponent();
        }

        private void btnCargar_Click(object sender, EventArgs e)
        {
            btnCargar.Enabled = false;
           

            if (fnImportarExcel.Importar("Hoja1", dgvData))
            {
                btnCargar.Enabled = true;
                btnProcesar.Enabled = true;
            }
        }

        private void btnProcesar_Click(object sender, EventArgs e)
        {
            DataTable investigacion = (DataTable)dgvData.DataSource;
            bool flag = fnScrap.InvestigacionTelefonos(investigacion, dgvData);

            if (flag)
            {
                btnProcesar.Enabled = false;
                btnExcel.Enabled = true;
            }
            
            
        }
    }
}
