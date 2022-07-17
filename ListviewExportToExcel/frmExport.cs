using System;
using System.Windows.Forms;

namespace ListviewExportToExcel{
    public partial class frmExport : Form {

        public frmExport(){
            InitializeComponent();
        }

        private void frmExport_Load(object sender, EventArgs e){
            //Exibe Detalhes
            lvlContatos.View = View.Details;
            //Permite selecionar a linha toda
            lvlContatos.FullRowSelect = true;
            //Exibe as linhas no listview
            lvlContatos.GridLines = true;
            //Permite que o usuario edite texto n listview
            lvlContatos.LabelEdit = true;
            //
            lvlContatos.Columns.Add("Nome", 300, HorizontalAlignment.Left);
            lvlContatos.Columns.Add("Celular", 150, HorizontalAlignment.Right);
            lvlContatos.Columns.Add("E-mail", 300, HorizontalAlignment.Left);
            //
            textNome.Select();

        }

        private void limpaCampos() {
            textNome.Text = string.Empty;
            textCelular.Text = string.Empty;
            textEmail.Text = string.Empty;
        }

        private void btnInserir_Click(object sender, EventArgs e){

            if(textNome.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("Você deve informar o nome!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (!textCelular.MaskCompleted){
                MessageBox.Show("Você deve informar o Telefone completo!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (textEmail.Text.Trim().Equals(string.Empty)){
                MessageBox.Show("Você deve informar o e-mail!", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //Definir os itens na lista
            ListViewItem list = new ListViewItem(textNome.Text.Trim());
            list.SubItems.Add(textCelular.Text);
            list.SubItems.Add(textEmail.Text.Trim());
            //Adicionar o item criado a listiew
            lvlContatos.Items.Add(list);


            this.limpaCampos();
        }

        private void btnLimpar_Click(object sender, EventArgs e){
            this.limpaCampos();
            lvlContatos.Items.Clear();
        }

        private void btnExportar_Click(object sender, EventArgs e){
            /*
             *
             *Adicionar em REFERENCIAS / Microsoft Excel 16.
             *
             */

            // QUANDO ADICIONAR AS REFERENCIAS, RETIRAR O COMENTARIO ABAIXO...

            
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            int linha = 2, coluna = 1;

            ws.Cells[1, 1] = lvlContatos.Columns[0].Text;
            ws.Cells[1, 2] = lvlContatos.Columns[1].Text;
            ws.Cells[1, 3] = lvlContatos.Columns[2].Text;

            foreach (ListViewItem list in lvlContatos.Items){

                coluna = 1;
                foreach (ListViewItem.ListViewSubItem item in list.SubItems){

                    ws.Cells[linha, coluna] = item.Text;
                    coluna++;
                }
                linha++;
            }
            
        }
    }
}
