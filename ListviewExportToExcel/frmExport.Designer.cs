
namespace ListviewExportToExcel
{
    partial class frmExport
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // frmExport
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "frmExport";
            this.Load += new System.EventHandler(this.frmExport_Load_1);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textNome;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MaskedTextBox textCelular;
        private System.Windows.Forms.TextBox textEmail;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnInserir;
        private System.Windows.Forms.ListView lvlContatos;
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.Button btnLimpar;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
    }
}

