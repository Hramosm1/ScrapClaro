using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ScrapClarov1.Funciones
{
    class fnImportarExcel
    {
        private static OleDbConnection conexionExcel(string ruta)
        {
            OleDbConnection conexion = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source='" + ruta + "';Extended Properties='Excel 12.0 Xml;HDR=Yes'");
            return conexion;
        }

        public static bool Importar(string nombreHoja, DataGridView grid) //Importar el excel a una data table
        {
            bool flag = false;
            DataTable dt = new DataTable();
            string ruta = "";
            OleDbConnection conexion;
            OleDbDataAdapter adaptador;

            try
            {
                OpenFileDialog Explorador = new OpenFileDialog();
                Explorador.Filter = "Excel files |*.xlsx";
                Explorador.Title = "Seleccionar archivo";

                if (Explorador.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    ruta = Explorador.FileName;
                }

                if (ruta != "")
                {
                    conexion = conexionExcel(ruta);
                    adaptador = new OleDbDataAdapter("select * from [" + nombreHoja + "$]", conexion);

                    adaptador.Fill(dt);
                    conexion.Close();

                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("Archivo cargado correctamente.", "Teléfonos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        grid.DataSource = dt;
                        flag = true;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ha ocurrido un error al momento de cargar el excel " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return flag;

        }
    }
}
