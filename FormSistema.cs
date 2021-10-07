using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace facturacion
{

    public partial class FormSistema : Form
    {
        
        public FormSistema()
        {
            InitializeComponent(); 
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnempresa_Click(object sender, EventArgs e)
        {
            FormFact mv = new FormFact();
            //this.Hide();
            mv.ShowDialog();
            //this.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FormGastos mv = new FormGastos();
            this.Hide();
            mv.ShowDialog();
            this.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            //Variable que abrira el archivo de access
            DialogResult abrirArchivoResultado = openFileDialog1.ShowDialog();
            if (abrirArchivoResultado == DialogResult.OK)
            {
                string RutaArhivo = openFileDialog1.FileName;
                FileInfo datoArchivo = new FileInfo(RutaArhivo);
                string archivoextension = datoArchivo.Extension;
                if (archivoextension == ".accdb")
                {
                    Entidades.Variablesglobales.rutaAccess = RutaArhivo;
                    //Variablesglobales.rutaAccess = RutaArhivo;
                    label5.Text = RutaArhivo;
                    //InciarDatos();
                }
            }
        }

        private void btnEmp_Click(object sender, EventArgs e)
        {
            FormDatoCliente mv = new FormDatoCliente();
            this.Hide();
            mv.ShowDialog();
            this.Show();
        }

        private void btnOferta_Click(object sender, EventArgs e)
        {
            FormOferta of = new FormOferta();
            this.Hide();
            of.ShowDialog();
            this.Show();
        }



    }
}
