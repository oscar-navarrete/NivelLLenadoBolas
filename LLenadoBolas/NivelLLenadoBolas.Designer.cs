using System.Configuration;

namespace LLenadoBolas
{
    partial class NivelLLenadoBolas
    {
        /// <summary> 
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary> 
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {

            string segundos = ConfigurationManager.AppSettings["Tiempo"];
            double tiempo = double.Parse(segundos);

            this.stLapso = new System.Timers.Timer();
            ((System.ComponentModel.ISupportInitialize)(this.stLapso)).BeginInit();
            // 
            // stLapso
            // 
            this.stLapso.Enabled = true;
            this.stLapso.Interval = tiempo;
            this.stLapso.Elapsed += new System.Timers.ElapsedEventHandler(this.stLapso_Elapsed);
            // 
            // NivelLLenadoBolas
            // 
            this.ServiceName = "NivelLLenadoBolas";
            ((System.ComponentModel.ISupportInitialize)(this.stLapso)).EndInit();

        }

        #endregion

        private System.Timers.Timer stLapso;
    }
}
