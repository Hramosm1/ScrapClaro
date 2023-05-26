using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Data;
using System.Windows.Forms;
using System.Threading;

namespace ScrapClarov1.Funciones
{
    class fnScrap
    {
        public static bool InvestigacionTelefonos(DataTable Investigacion, DataGridView grid, IWebDriver driver)
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

            foreach (DataRow investigar in Investigacion.Rows)
            {
                cont += 1;

                try
                {
                    string telefono = investigar.ItemArray[0].ToString();

                    IWebElement texTelefono = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/input"));
                    texTelefono.Clear();
                    texTelefono.SendKeys(telefono);
                    Thread.Sleep(3000);

                    if (driver.PageSource.Contains("Mensajes [226203647]"))
                    {
                        IWebElement error3 = driver.FindElement(By.XPath("/html/body/div[18]/div[2]/div[1]/div/div/div[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/em/button"));
                        error3.Click();
                    }

                    IWebElement buscar = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/table/tbody/tr[2]/td[2]/em/button"));
                    buscar.Click();
                    Thread.Sleep(10000);

                    if (driver.PageSource.Contains("Mensajes [226194535]"))
                    {
                        IWebElement error1 = driver.FindElement(By.XPath("/html/body/div[18]/div[2]/div[1]/div/div/div[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/em/button"));
                        error1.Click();
                    }

                    if (driver.PageSource.Contains("Nuevo Cliente"))
                    {
                        IWebElement primerCliente = driver.FindElement(By.XPath("/html/body/div[18]/div[2]/div[1]/div/div/div/div/div/div[1]/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div[1]/div[2]/div/div[2]/table/tbody/tr/td[4]/div/table/tbody/tr/td"));
                        primerCliente.Click();
                        IWebElement error2 = driver.FindElement(By.XPath("/html/body/div[18]/div[2]/div[1]/div/div/div/div/div/div[2]/div/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/em/button"));
                        error2.Click();
                        Thread.Sleep(5000);
                    }

                    var nombres = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[2]")).Text;
                    var apellidos = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[4]")).Text;
                    var nit = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[2]/td[2]")).Text;
                    var dpi = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[2]/td[4]")).Text;
                    var codCliente = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[5]/td[2]")).Text;
                    var fechaNac = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[5]/td[4]")).Text;
                    var cliente = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[6]/td[4]")).Text;
                    var email = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[1]/td[6]")).Text;
                    var representante = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[5]/td[6]")).Text;
                    var telefono1 = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[2]/td[6]")).Text;
                    var telefono2 = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[3]/td[6]")).Text;
                    var direccion = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[3]/td[2]")).Text;
                    var direccion2 = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[3]/td[4]")).Text;
                    var departamento = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[4]/td[4]")).Text;
                    var canton = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[4]/td[6]")).Text;
                    var onBase = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/div/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[1]/div/div/div/div/div/div/div/div/div/div[2]/div/div/div/div/div/div/div/div/div/table/tbody/tr[6]/td[2]")).Text;

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

            IWebElement config = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[1]/div/div/div/div[4]/div/div/div/div[2]/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/em/button"));
            config.Click();

            IWebElement salir = driver.FindElement(By.XPath("/html/body/div[16]/ul/li[1]/a"));
            salir.Click();

            IWebElement confirmar = driver.FindElement(By.XPath("/html/body/div[20]/div[2]/div[2]/div/div/div/div[1]/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]"));
            confirmar.Click();

            driver.Close();
            driver.Quit();

            grid.DataSource = resultados;
            return flag;
        }

        public static bool inicializacionDriver(DataTable Investigacion, DataGridView grid)
        {
            bool flag = false;
            DataTable dt = new DataTable();
            DataTable resultados = Investigacion;

            string sitio = "http://172.17.224.54/Beesion.CrmAmxCenam.Gt/#";
            string user = "CLAROGT\\BRAYAN.PEREZ";
            string pass = "Mayo+2023";
            try
            {
                // Inicializar el controlador de Selenium
                ChromeOptions options = new ChromeOptions();
                options.AddArgument("--start-maximized");

                IWebDriver driver = new ChromeDriver(options);

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

                InvestigacionTelefonos(resultados, grid, driver);

                
            }
            catch (Exception ex)
            {
               
            }
            return flag;
        }

    }
}
