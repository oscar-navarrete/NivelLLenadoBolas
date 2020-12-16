using Microsoft.SolverFoundation.Services;
using Microsoft.SolverFoundation.Solvers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using OSIsoft.AF;
using System.Net;
using OSIsoft.AF.Time;
using OSIsoft.AF.Search;
using OSIsoft.AF.Asset;
using OSIsoft.AF.PI;

namespace LLenadoBolas
{
    partial class NivelLLenadoBolas : ServiceBase
    {

        bool blBandera = false;
       
        public NivelLLenadoBolas()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: agregar código aquí para iniciar el servicio.
            stLapso.Start();
        }

        protected override void OnStop()
        {
            // TODO: agregar código aquí para realizar cualquier anulación necesaria para detener el servicio.
            stLapso.Stop();
        }

        private void stLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (blBandera) return;

            try
            {
                blBandera = true;

                EventLog.WriteEntry("Se inicio proceso de Calculo Nivel LLenado de Bolas", EventLogEntryType.Information);
                 CalculosLLenadoBolas();
            }
            catch (Exception ex)
            {
                //oLog.Add(ex.Message);
                EventLog.WriteEntry(ex.Message, EventLogEntryType.Error);
            }

            blBandera = false;
        }

 
        public static void CalculosLLenadoBolas()
        {
            //string path = System.IO.Directory.GetCurrentDirectory();
            string path = @"C:\NivelBolas\";
            Log oLog = new Log(path);

            string servidor = ConfigurationManager.AppSettings["ServidorAF"];
            string usuario = ConfigurationManager.AppSettings["Usuario"];
            string password = ConfigurationManager.AppSettings["Password"];
            string dominio = ConfigurationManager.AppSettings["Dominio"];
            string bd = ConfigurationManager.AppSettings["BD"];
            string modelo = ConfigurationManager.AppSettings["Modelo"];
            string servidorpi = ConfigurationManager.AppSettings["ServidorPI"];

            AFElement model = new AFElement();
            PISystems AF = new PISystems();
            PISystem AFSrv = AF[servidor];
            string AFConnectionStatus = string.Empty;
            string AFConnectionStatusError = string.Empty;

            try
            {
                NetworkCredential credential = new NetworkCredential(usuario, password, dominio);
                AFSrv.Connect(credential);
                AFConnectionStatus = "Good";
                AFConnectionStatusError = " Conexión Exitosa a Servidor AF " + AFSrv.Name;
                oLog.Add(AFConnectionStatusError);
                
            }
            catch (Exception ex)
            {
                // Expected exception since credential needs a valid user name and password.
                AFConnectionStatus = "Failed";
                AFConnectionStatusError =  " Servidor AF no existe o Revisar Credenciales en Archivo .ini" + " MsgError:" + ex.Message;
                oLog.Add(AFConnectionStatusError);
                oLog.Add("Fin de Ciclo");
                AFSrv.Disconnect();
                AFSrv.Dispose();
                //Environment.Exit(0);
            }

            //Elegir Base de Datos AF
            string AFConnectionStatusBD = string.Empty;
            string AFConnectionStatusBDError = string.Empty;
            string templateName = "Molinos";
            if (AFConnectionStatus == "Good")
            {
                AFDatabases AFSrvBDs = AFSrv.Databases;
                AFDatabase AFSrvBD = AFSrvBDs[bd];
                
                oLog.Add("---Inicio Proceso---");

                using (AFElementSearch elementQuery = new AFElementSearch(AFSrvBD, "TemplateSearch", string.Format("template:\"{0}\"", templateName)))
                {
                    //elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);
                    foreach (AFElement element in elementQuery.FindElements())
                    {

                        if (element.Attributes["Habilitado"].GetValue().Value.ToString()=="True")
                        {
                            double mill_speed = 0.0;
                            double balls_filling = 0.0;
                            double velocidad = element.Attributes["Velocidad"].GetValue().ValueAsDouble();
                            double diametro = element.Attributes["Diametro"].GetValue().ValueAsDouble();
                            double angulo = element.Attributes["Angulo"].GetValue().ValueAsDouble();
                            double densidad_bolas = element.Attributes["Densidad_Bolas"].GetValue().ValueAsDouble();
                            double ore_density = element.Attributes["Densidad_Mineral"].GetValue().ValueAsDouble();
                            double largo = element.Attributes["Largo"].GetValue().ValueAsDouble();
                            double intersticios = element.Attributes["LLenado_Intersticios"].GetValue().ValueAsDouble();
                            double losses = element.Attributes["Perdidas_Energia"].GetValue().ValueAsDouble();
                            double slurry_density = element.Attributes["Slurry_Density"].GetValue().ValueAsDouble();
                            double solidos = element.Attributes["Solidos_Molinos"].GetValue().ValueAsDouble();

                            //Código de Prueba
                            //mill_speed = element.Attributes["Mill_Speed"].GetValue().ValueAsDouble();

                            mill_speed = Solver(velocidad, diametro);
                            element.Attributes["Mill_Speed"].SetValue(new AFValue(mill_speed));

                            double potencia_prom = element.Attributes["Potencia_Prom"].GetValue().ValueAsDouble();
                            balls_filling = Solver_Potencia(Math.Round(potencia_prom, 0), diametro, largo, Math.Round(mill_speed, 0), densidad_bolas, Math.Round(solidos, 2), ore_density, angulo, losses, intersticios);

                            element.Attributes["Balls_Filling"].SetValue(new AFValue(balls_filling));

                            oLog.Add("Molino: " + element.Name + " Mill_Speed: " + mill_speed + " Balls_Filling: " + balls_filling);

                        }

                    }

                    oLog.Add("---Fin Proceso---");
                    AFSrv.Disconnect();
                    AFSrv.Dispose();

                }


               
            }
        }

        public static double Solver(double rpm, double diametro)
        {
            SolverContext context = SolverContext.GetContext();
            context.ClearModel();
            Model model = context.CreateModel();

            Decision x = new Decision(Domain.RealNonnegative, "Mill_Speed");

            model.AddDecisions(x);
            model.AddConstraints("one", rpm == (76.6 / Math.Pow(diametro, 0.5)) * (x / (double)100));

            SimplexDirective simplex = new SimplexDirective();

            // Solve the problem
            context.Solve(simplex);
            
            return x.GetDouble();
            
        }

        public static double Solver_Potencia(double potencia, double diametro, double largo, double millSpeed, double densidad_bolas, double solidos, double densidad_mineral, double angulo, double losses, double inter)
        {
       
            SolverContext context1 = SolverContext.GetContext();
            context1.ClearModel();
            Model model1 = context1.CreateModel();

            Decision y = new Decision(Domain.RealNonnegative, "Balls_Filling");
            
            model1.AddDecisions(y);

            model1.AddConstraints("one", potencia == (0.238 * Math.Pow(diametro, 3.5) * (largo / (double)diametro) * (millSpeed / (double)100) * ((((1 - 0.4) * densidad_bolas * (y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4) + ((1 / (double)((solidos / (double)100) / densidad_mineral + (1 - solidos / (double)100))) * (y / (double)100 - y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4) + ((1 / (double)((solidos / (double)100) / densidad_mineral + (1 - solidos / (double)100))) * (inter / (double)100) * 0.4 * (y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4)) / ((y / (double)100) * Math.PI * Math.Pow((diametro * 0.305), 2) * (largo * 0.305) / (double)4)) * (y / (double)100 - 1.065 * y * y / (double)10000) * Math.Sin(angulo * Math.PI / 180)) / (1 - losses / (double)100));

            SimplexDirective simplex = new SimplexDirective();
        
            Directive n = new Directive();
            n.TimeLimit = 300000;
            n.WaitLimit = 420000;
            //simplex.IterationLimit = 100;

            // Solve the problem
            Solution sol = context1.Solve(n);
            
            ///Report report = sol.GetReport();
            
            return y.GetDouble();

        }

       

    }
}
