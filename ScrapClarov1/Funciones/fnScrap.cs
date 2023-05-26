using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data;
using System.Windows.Forms;

namespace ScrapClarov1.Funciones
{
    class fnScrap
    {
        public static bool InvestigacionTelefonos(DataTable Investigacion, DataGridView grid)
        {
            bool flag = false;
            DataTable dt = new DataTable();
            DataTable resultados = Investigacion;
            int cont = -1;

            resultados.Columns.Add("Nombres");
            resultados.Columns.Add("Apellidos");
            resultados.Columns.Add("NIT");
            resultados.Columns.Add("DPI");
            resultados.Columns.Add("Codigo Cliente");
            resultados.Columns.Add("Fecha Nacimiento");
            resultados.Columns.Add("Cliente");
            resultados.Columns.Add("Email");
            resultados.Columns.Add("Representante");
            resultados.Columns.Add("Telefono 1");
            resultados.Columns.Add("Telefono 2");
            resultados.Columns.Add("Direccion");
            resultados.Columns.Add("Direccion Adi");
            resultados.Columns.Add("Departamento");
            resultados.Columns.Add("Canton");
            resultados.Columns.Add("onBase");

            //if (inicializacionDriver())
            //{
                foreach (DataRow investigar in Investigacion.Rows)
                {
                    cont += 1;

                    try
                    {
                        var nombres = "Nombres " + cont;
                        var apellidos = "Apellidos " + cont;
                        var nit = "nit " + cont;
                        var dpi = "dpi " + cont;
                        var codCliente = "codCliente " + cont;
                        var fechaNac = "fecha Nac " + cont;
                        var cliente = "cliente " + cont;
                        var email = "email " + cont;
                        var representante = "representante " + cont;
                        var telefono1 = "telefono1 " + cont;
                        var telefono2 = "telefono2 " + cont;
                        var direccion = "direccion " + cont;
                        var direccion2 = "direccion 2" + cont;
                        var departamento = "departamento " + cont;
                        var canton = "canton " + cont;
                        var onBase = "onBase " + cont;

                        resultados.Rows[cont]["Nombres"] = nombres; //Ingreso la respuesta a la grid con la data
                        resultados.Rows[cont]["Apellidos"] = apellidos;
                        resultados.Rows[cont]["NIT"] = nit;
                        resultados.Rows[cont]["DPI"] = dpi;
                        resultados.Rows[cont]["Codigo Cliente"] = codCliente;
                        resultados.Rows[cont]["Fecha Nacimiento"] = fechaNac;
                        resultados.Rows[cont]["Cliente"] = cliente;
                        resultados.Rows[cont]["Email"] = email;
                        resultados.Rows[cont]["Representante"] = representante;
                        resultados.Rows[cont]["Telefono 1"] = telefono1;
                        resultados.Rows[cont]["Telefono 2"] = telefono2;
                        resultados.Rows[cont]["Direccion"] = direccion;
                        resultados.Rows[cont]["Direccion Adi"] = direccion2;
                        resultados.Rows[cont]["Departamento"] = departamento;
                        resultados.Rows[cont]["Canton"] = canton;
                        resultados.Rows[cont]["onBase"] = onBase;

                        flag = true;

                    }
                    catch (Exception ex)
                    {

                    }

                }
            //}

            grid.DataSource = resultados;
            return flag;
        }

        public static bool inicializacionDriver()
        {
            bool flag = false;
            string sitio = "http://172.17.224.54/Beesion.CrmAmxCenam.Gt/#";
            string user = "CLAROGT\\BRAYAN.PEREZ";
            string pass = "Mayo+2023";
            try
            {
                // Inicializar el controlador de Selenium
                IWebDriver driver = new ChromeDriver();

                //Abrir el sitio Web
                driver.Navigate().GoToUrl(sitio);

                //Esperar a que se abra la aplicacion
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                //Ingresar usuario y contraseña
                IWebElement usuario = driver.FindElement(By.Id("txtAccount"));
                IWebElement password = driver.FindElement(By.Id("txtPassword"));
                IWebElement botonIngresar = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/table/tbody/tr[2]/td[2]/em/button"));
                usuario.SendKeys(user);
                password.SendKeys(pass);
                botonIngresar.Click();

                IWebElement pantallaUnica = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[1]/div[2]/div/div[2]/div[1]"));
                pantallaUnica.Click();

                flag = true;
            }
            catch (Exception ex)
            {
               
            }

            return flag;
        }

    }
}
