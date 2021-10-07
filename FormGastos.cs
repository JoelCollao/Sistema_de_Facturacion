using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static facturacion.Entidades;

namespace facturacion
{
    public partial class FormGastos : Form
    {

        String rutaAcces = Entidades.Variablesglobales.rutaAccess;
        String microsolftAccessDatabaseConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = ";

        String selectDataFromMSAccessDatabaseQuery = "SELECT * FROM TablaGasto ";
        //CONSULTA PARA TRAER DATOS DEL TABLA DE CONCEPTO GASTO
        String selectDataConceptoDatabaseQuery = "SELECT ConceptoGasto FROM TablaConceptoGastos";

        //TRAER EL VALOR DE RUC DE TABLA DE CLIENTES
        String QueryDatosTablaCliente = "SELECT * FROM TablaCliente";

        //TRAER LA TABLA DE CLIENTES
        String QueryDatosTablaFactura = "SELECT * FROM TablaFactura";

        String InsertDataIntoMicrosoftAccessDatabase = "INSERT INTO TablaGasto (nomb_emp,nomb_prove,ruc_pr,proyecto,tipo,fechaRecibo,Mes_Reci,Anho_Reci,numFactClient, ruc_cl, costo, ConIgv, concepto, observaciones ,PeriImputaci,quienPag,fechaPagEmpresa,Mes_Emp,Anho_Emp,NumFactProve) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

        String UpdateMicrosoftAccessDataQuery = "UPDATE TablaGasto SET nomb_emp= ?,nomb_prove= ?,ruc_pr= ?,proyecto= ?,tipo= ?,fechaRecibo= ?,Mes_Reci= ?,Anho_Reci= ?,numFactClient= ?,ruc_cl= ?, costo= ?, ConIgv= ?,concepto= ?,observaciones= ?,PeriImputaci= ?,quienPag= ?,fechaPagEmpresa= ?,Mes_Emp= ?,Anho_Emp= ?,NumFactProve =? WHERE Id = ?";
        String DeleteMicrosoftAccessDataQuery = "DELETE * FROM TablaGasto WHERE Id = ?";


        OleDbConnection AccessDatabaseConecctionStringOleConecction = null;
        List<TabFactura> ListaFactura = new List<TabFactura>();
        List<TablaGasto> ListaTablaGastos = new List<TablaGasto>();
        List<TabClientes> ListaRUC = new List<TabClientes>();
        List<TablaConceptGasto> ListaConcepto = new List<TablaConceptGasto>();

        public FormGastos()
        {
            InitializeComponent();
        }


        public void OpenConexion()
        {
            AccessDatabaseConecctionStringOleConecction = new OleDbConnection(microsolftAccessDatabaseConnectionString + rutaAcces);
            if (AccessDatabaseConecctionStringOleConecction.State == ConnectionState.Closed)
            {
                AccessDatabaseConecctionStringOleConecction.Open();
            }
        }

        public void CloseConexion()
        {
            if (AccessDatabaseConecctionStringOleConecction.State == ConnectionState.Open)
            {
                AccessDatabaseConecctionStringOleConecction.Close();
            }
        }

        public void llenardatosConceptoGastos()
        {
            OpenConexion();
            ListaConcepto = PopulatedComboConcepto();
            cbConcepto.DisplayMember = "ConceptoGasto";
            cbConcepto.ValueMember = "ConceptoGasto";


            OpenConexion();
            ListaRUC = PopulatedAddComboboxTablaClienteRUC();
            cbRUC.DisplayMember = "ruc_cl";
            cbRUC.ValueMember = "ruc_cl";

            OpenConexion();
            ListaFactura = PupulatedAddTablaFactura();
            cbNumFactura.DisplayMember = "numFactura";
            cbNumFactura.ValueMember = "numFactura";
        }

        private void IniciarDatos()
        {
            OpenConexion();
            ListaTablaGastos = PopulatedAddTablaGastos();


            llenardatosConceptoGastos();
            dataGridView1.DataSource = ListaTablaGastos;

            //PARA LLENAR LISTA DE CONCEPTO DE GASTO
            cbConcepto.DataSource = ListaConcepto;
            
          
            TabClientes objetoSeleccionar = new TabClientes();
            objetoSeleccionar.ruc_cl = "Seleccionar";
            objetoSeleccionar.nomb_emp = "";
            ListaRUC.Add(objetoSeleccionar);
            int indexvalores = ListaRUC.Count;
            cbRUC.DataSource = ListaRUC;
            cbRUC.SelectedIndex = indexvalores - 1;

            //NOSE CARGA 2 VECES EN COMBOBOX --- RECORDAR ESTO NO VA
            //cbNumFactura.DataSource = ListaFactura;

            TabFactura objetoSelect = new TabFactura();
            objetoSelect.numFactura = "Seleccione #Factura";
            objetoSelect.import = 0;
            ListaFactura.Add(objetoSelect);
            int indexvalor = ListaFactura.Count;
            cbNumFactura.DataSource = ListaFactura;
            cbNumFactura.SelectedIndex = indexvalor - 1;

        }


       

        //SELECCIONAR UN DATO DEL COMBOBOX
      

        //LLENADO EN COMBOBOZ DE LOS DATOS DE PROYECTO DE LA TABLA DE FACTURA
        public List<TabClientes> PopulatedAddComboboxTablaClienteRUC()
        {
            List<TabClientes> ListaRUC = new List<TabClientes>();
            try
            {
                if (cbRUC.Items.Count > 0)
                {
                    cbRUC.DataSource = null;
                }

                OleDbCommand QueryextracionDedatosCombobox = new OleDbCommand(QueryDatosTablaCliente, AccessDatabaseConecctionStringOleConecction);
                OpenConexion();

                TabClientes objeto;
                OleDbDataReader reader = QueryextracionDedatosCombobox.ExecuteReader();
                
                while (reader.Read())
                {
                    objeto = new TabClientes();
                    objeto.ruc_cl = reader[1].ToString();
                    objeto.nomb_emp = reader[2].ToString();

                    ListaRUC.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaRUC;
            }
            finally
            {
                CloseConexion();
            }
            return ListaRUC;
        }


        //LLENAR DATOS A UN COMBOBOX , USANDO EL READER Y CONEXION A BD ACCEES
        public List<TablaConceptGasto> PopulatedComboConcepto()
        {
            //control k + control s  : sacar el TRY CATH
            //creamos la lista en una instania de lista de la tabla Factura
            List<TablaConceptGasto> listaConcepto = new List<TablaConceptGasto>();
            try
            {
                //// borrar filas de cuadrícula de datos antes de cargar datos de accesos de microsoft
                if (cbConcepto.Items.Count > 0)
                {
                    cbConcepto.DataSource = null;
                }
                OleDbCommand conceptoMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(selectDataConceptoDatabaseQuery, AccessDatabaseConecctionStringOleConecction);
                //abrimos la conexion a base de datos
                OpenConexion();

                //Es el nombre de la variable , se crea para poder instanciar y poder llenar los datos y
                //para que se guarden de forma ordenada se crea un lista
                TablaConceptGasto objeto;

                OleDbDataReader reader = conceptoMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                // bucle para leer los datos de Microsoft Access Database
                while (reader.Read())
                {
                    objeto = new TablaConceptGasto();

                    objeto.ConceptoGasto = reader[0].ToString();

                    listaConcepto.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return listaConcepto;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseConexion();
            }
            return listaConcepto;
        }

        public List<TabFactura> PupulatedAddTablaFactura()
        {
            List<TabFactura> listFactura = new List<TabFactura>();
            try
            {
                //// borrar filas de cuadrícula de datos antes de cargar datos de accesos de microsoft
                if (cbNumFactura.Items.Count > 0)
                {
                    cbNumFactura.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryDatosTablaFactura, AccessDatabaseConecctionStringOleConecction);
                //abrimos la conexion a base de datos
                OpenConexion();

                //Es el nombre de la variable , se crea para poder instanciar y poder llenar los datos y
                //para que se guarden de forma ordenada se crea un lista
                TabFactura objeto;

                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                // bucle para leer los datos de Microsoft Access Database
                while (reader.Read())
                {
                    objeto = new TabFactura();

                    objeto.numFactura = reader[4].ToString();
                    if (Information.IsDBNull(reader[6]) != true)
                    {
                        objeto.import = float.Parse(reader[6].ToString());
                    }

                    listFactura.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return listFactura;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseConexion();
            }
            return listFactura;
        }


        public List<TablaGasto> PopulatedAddTablaGastos()
        {
            //control k + control s  : sacar el TRY CATH
            //creamos la lista en una instania de lista de la tabla Factura
            List<TablaGasto> listaGastos = new List<TablaGasto>();
            int IDError = 0;
            int ColumnaError = 0;
            try
            {
                //// borrar filas de cuadrícula de datos antes de cargar datos de accesos de microsoft
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.DataSource = null;
                }
                OleDbCommand populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(selectDataFromMSAccessDatabaseQuery, AccessDatabaseConecctionStringOleConecction);
                //abrimos la conexion a base de datos
                OpenConexion();
                //Es el nombre de la variable , se crea para poder instanciar y poder llenar los datos y
                //para que se guarden de forma ordenada se crea un lista
                TablaGasto objeto;
                OleDbDataReader reader = populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                // bucle para leer los datos de Microsoft Access Database
                while (reader.Read())
                {
                    objeto = new TablaGasto();
                    ColumnaError = 1;
                    objeto.Id = (int)reader[0];
                    IDError = (int)reader[0];
                    ColumnaError = 2;
                    objeto.nomb_emp = reader[1].ToString();
                    ColumnaError = 3;
                    objeto.nomb_prove = reader[2].ToString();
                    ColumnaError = 4;
                    objeto.ruc_pr = reader[3].ToString();
                    ColumnaError = 5;
                    objeto.proyecto = reader[4].ToString();
                    ColumnaError = 6;
                    objeto.tipo = reader[5].ToString();

                    //---------------------------------------

                    //valores en fechas
                    ColumnaError = 7;
                    if (Information.IsDBNull(reader[6]) != true)
                    {
                        objeto.fechaRecibo = (DateTime)reader[6];
                    }
                    ColumnaError = 8;
                    objeto.Mes_Reci = reader[7].ToString();
                    ColumnaError = 9;
                    objeto.Anho_Reci = reader[8].ToString();
                    ColumnaError = 10;
                    objeto.numFactClient = reader[9].ToString();
                    ColumnaError = 11;
                    objeto.ruc_cl = reader[10].ToString();
                    ColumnaError = 12;
                    if (Information.IsDBNull(reader[11]) != true)
                    {
                        objeto.costo = float.Parse(reader[11].ToString());
                    }
                    ColumnaError = 13;

                    if (Information.IsDBNull(reader[12]) != true)
                    {
                        objeto.ConIgv = float.Parse(reader[12].ToString());
                        ColumnaError = 14;
                    }
                    
                    objeto.concepto = reader[13].ToString();
                    ColumnaError = 15;
                    objeto.observaciones = reader[14].ToString();
                    ColumnaError = 16;
                    objeto.PeriImputaci = reader[15].ToString();
                    ColumnaError = 17;
                    objeto.quienPag = reader[16].ToString();

                    //---------------------------------------------------------
                    ColumnaError = 18;
                    if (Information.IsDBNull(reader[17]) != true)
                    {
                        objeto.fechaPagEmpresa = (DateTime)reader[17];
                    }
                    ColumnaError = 19;
                    objeto.Mes_Emp = reader[18].ToString();
                    ColumnaError = 20;
                    objeto.Anho_Emp = reader[19].ToString();
                    objeto.NumFactProve = reader[20].ToString();

                        listaGastos.Add(objeto);
                    }
                    reader.Close();
                }
            catch (Exception ex)
            {

                //MessageBox.Show(ex.Message);
                string mensajeErrror = "El codigo con : " + IDError.ToString() + " Tiene un error. La columna es: " + ColumnaError.ToString();

                MessageBox.Show(ex.Message + " " + mensajeErrror);
                return listaGastos;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseConexion();
            }
            return listaGastos;
        }   

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            limpiar();

            txtID.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtNombEmpre.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtNomProveedor.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            
            txtRUC_PR.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtProject.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            cbTipoFactura.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();

            //---------------------------------------------------------------

            if (dataGridView1.CurrentRow.Cells[6].Value != null)
            {
                txtFechaRecibo.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            }
                
            txtMes_Recibo.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            txtAnho_Recibo.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            cbNumFactura.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();

            //----------------------------------------------------------------

            cbRUC.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            txtCosto.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            txtConIgv.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            cbConcepto.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            txtObserva.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            txtPeriodImput.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            txtQuienPag.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();

            if (dataGridView1.CurrentRow.Cells[17].Value !=null)
            {
                txtFechaEmpresa.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            }
            

            txtMes_Empresa.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            txtAnho_Empresa.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            txtNroFactProvee.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();

        }

        private void limpiar()
        {

            txtID.Text = String.Empty;
            txtNombEmpre.Text = String.Empty;
            txtNomProveedor.Text = String.Empty;
            txtRUC_PR.Text = String.Empty;
            txtProject.Text = String.Empty;
            cbNumFactura.Text = String.Empty;
            txtFechaRecibo.Text = String.Empty;
            txtMes_Recibo.Text = String.Empty;
            txtAnho_Recibo.Text = String.Empty;
            cbTipoFactura.Text = String.Empty;
            txtNroFactProvee.Text = String.Empty;
            cbRUC.Text = String.Empty;
            txtCosto.Text = String.Empty;
            txtConIgv.Text = String.Empty;
            cbConcepto.Text = String.Empty;
            txtObserva.Text = String.Empty;
            txtPeriodImput.Text = String.Empty;
            txtQuienPag.Text = String.Empty;
            txtFechaEmpresa.Text = String.Empty;
            txtMes_Empresa.Text = String.Empty;
            txtAnho_Empresa.Text = String.Empty;
        }


        private void txtCosto_TextChanged(object sender, EventArgs e)
        {
            txtConIgv.Text = string.Empty;
        }


        void CargaInicio()
        {

            try
            {
                String rutaArchivo = Entidades.Variablesglobales.rutaAccess;

                FileInfo datoArchivo = new FileInfo(rutaArchivo);
                String archivoExtension = datoArchivo.Extension;

                if (archivoExtension == ".accdb")
                {
                    rutaAcces = rutaArchivo;
                    lblRuta.Text = rutaArchivo;
                    IniciarDatos();
                    //MessageBox.Show("Archivo correcto");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Primero Necesita Realizar la Conexion a la Base de datos", "Verificar Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();

            }

        }

        private void FormGastos_Load(object sender, EventArgs e)
        {
            CargaInicio();
        }

        
        private void cbNumFactura_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //string numFactura = cbNumFactura.SelectedValue.ToString();

            //List<TabFactura> ListaRUCResultado = new List<TabFactura>();

            //ListaRUCResultado = ListaFactura.Where(x => x.numFactura == numFactura).ToList();

            //if (ListaRUCResultado.Count > 0)
            //{
            //    txtImporte.Text = ListaRUCResultado[0].import.ToString();
                
            //}
        }

        private void btnErase_Click(object sender, EventArgs e)
        {
            OpenConexion();

            try
            {
                OleDbCommand DeleteMicrosoftAccessDataQueryOleDbComand = new OleDbCommand(DeleteMicrosoftAccessDataQuery, AccessDatabaseConecctionStringOleConecction);
                DeleteMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Id", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);

                int DeleteMicrosoftAccessData = DeleteMicrosoftAccessDataQueryOleDbComand.ExecuteNonQuery();
                if (DeleteMicrosoftAccessData > 0)
                {
                    MessageBox.Show("Registro Borrado");
                    limpiar();
                    ListaTablaGastos = PopulatedAddTablaGastos();
                    dataGridView1.DataSource = ListaTablaGastos;

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                CloseConexion();
            }
        }

        private void btnExpo_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectAll();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    String elvalor = "";

                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        elvalor = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        DateTime datoFecha;
                        if (DateTime.TryParse(elvalor, out datoFecha))
                        {
                            if (dataGridView1.Rows[i].Cells[j].OwningColumn.DataPropertyName.ToUpper().Contains("FECHA"))
                            {
                                elvalor = datoFecha.ToString("dd/MM/yyyy");
                            }

                        }
                    }

                    xlWorkSheet.Cells[i + 2, j + 1] = elvalor;
                }
            }
            // save the application  
            xlWorkBook.SaveAs("C:\\Datos\\output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Exit from the application 
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            OpenConexion();
            OleDbCommand UpdateMicrosoftAccessDataQueryOleDbComand = new OleDbCommand(UpdateMicrosoftAccessDataQuery, AccessDatabaseConecctionStringOleConecction);
            try
            {
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmpre.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("nomb_prove", OleDbType.VarChar).Value = txtNomProveedor.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("ruc_pr", OleDbType.VarChar).Value = txtRUC_PR.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtProject.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("tipo", OleDbType.VarChar).Value = cbTipoFactura.Text;

                DateTime FechaReci;
                var fechaReciNull = (object)DBNull.Value;
                if (txtFechaRecibo.Text != "")
                {
                    if (DateTime.TryParse(txtFechaRecibo.Text, out FechaReci))
                    {
                        fechaReciNull = FechaReci.ToString("dd/MM/yyyy");
                    }
                }
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("fechaRecibo", OleDbType.VarChar).Value = fechaReciNull;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Mes_Reci", OleDbType.VarChar).Value = txtMes_Recibo.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Anho_Reci", OleDbType.VarChar).Value = txtAnho_Recibo.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("numFactClient", OleDbType.VarChar).Value = cbNumFactura.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = cbRUC.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("costo", OleDbType.VarChar).Value = txtCosto.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("ConIgv", OleDbType.VarChar).Value = txtConIgv.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("concepto", OleDbType.VarChar).Value = cbConcepto.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("observaciones", OleDbType.VarChar).Value = txtObserva.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("PeriImputaci", OleDbType.VarChar).Value = txtPeriodImput.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("quienPag", OleDbType.VarChar).Value = txtQuienPag.Text;

                DateTime fechaEmpresa;
                var fechaEmpresaNull = (object)DBNull.Value;
                if (txtFechaEmpresa.Text != "")
                {
                    if (DateTime.TryParse(txtFechaEmpresa.Text, out fechaEmpresa))
                    {
                        fechaEmpresaNull = fechaEmpresa.ToString("dd/MM/yyyy");
                    }
                }

                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("fechaPagEmpresa", OleDbType.VarChar).Value = fechaEmpresaNull;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Mes_Emp", OleDbType.VarChar).Value = txtMes_Empresa.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Anho_Emp", OleDbType.VarChar).Value = txtAnho_Empresa.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("NumFactProve", OleDbType.VarChar).Value = txtNroFactProvee.Text;
                UpdateMicrosoftAccessDataQueryOleDbComand.Parameters.AddWithValue("Id", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);

                int UpdateInsert = UpdateMicrosoftAccessDataQueryOleDbComand.ExecuteNonQuery();
                if (UpdateInsert > 0)
                {
                    MessageBox.Show("Registro Actualizado", "Cuadro de Verificación", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ListaTablaGastos = PopulatedAddTablaGastos();
                    dataGridView1.DataSource = ListaTablaGastos;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                CloseConexion();
            }
        }

        private void btnAgre_Click(object sender, EventArgs e)
        {

            OpenConexion();
            OleDbCommand insertDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(InsertDataIntoMicrosoftAccessDatabase, AccessDatabaseConecctionStringOleConecction);

            try
            {
                //los nombres de los campos debe ser correspondientes con la base de datos Access
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmpre.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_prove", OleDbType.VarChar).Value = txtNomProveedor.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_pr", OleDbType.VarChar).Value = txtRUC_PR.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtProject.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("tipo", OleDbType.VarChar).Value = cbTipoFactura.Text;
                //-----------------------------------------------------------------------------------------------------------------------------------------

                //Fecha de Recibo
                DateTime FechaReci;
                var FechaRecibidoNull = (object)DBNull.Value;
                if (txtFechaRecibo.Text != "")
                {
                    if (DateTime.TryParse(txtFechaRecibo.Text, out FechaReci))
                    {
                        FechaRecibidoNull = FechaReci.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaRecibo", OleDbType.VarChar).Value = FechaRecibidoNull;

                //------------------------------------------------------------------------------------------------------------------------------------------
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Mes_Reci", OleDbType.VarChar).Value = txtMes_Recibo.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Anho_Reci", OleDbType.VarChar).Value = txtAnho_Recibo.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("numFactClient", OleDbType.VarChar).Value = cbNumFactura.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = cbRUC.Text;

                if (Information.IsDBNull(txtCosto.Text) != true)
                {
                    insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("costo", OleDbType.VarChar).Value = txtCosto.Text;
                }
              

                //-------------------------------------------------------------------------------------------------------------------------------------------
                double calculoIgv;
                calculoIgv = (Convert.ToDouble(txtCosto.Text) * 0.18) + Convert.ToDouble(txtCosto.Text);
                txtConIgv.Text = System.Convert.ToString(calculoIgv);

                if (Information.IsDBNull(txtConIgv.Text) != true)
                {
                    insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ConIgv", OleDbType.VarChar).Value = txtConIgv.Text;
                }
               
                
                
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("concepto", OleDbType.VarChar).Value = cbConcepto.Text;
                
                
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("observaciones", OleDbType.VarChar).Value = txtObserva.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("PeriImputaci", OleDbType.VarChar).Value = txtPeriodImput.Text.ToUpper();
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("quienPag", OleDbType.VarChar).Value = txtQuienPag.Text.ToUpper();

                //----------------------------------------------------------------------------------------------------------------------------------------

                // FECHA DE LA EMPRESA
                DateTime FechaEmpre;
                var FechaEmpreNull = (object)DBNull.Value;
                if (txtFechaEmpresa.Text != "")
                {
                    if (DateTime.TryParse(txtFechaEmpresa.Text, out FechaEmpre))
                    {
                        FechaEmpreNull = FechaEmpre.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaPagEmpresa", OleDbType.VarChar).Value = FechaEmpreNull;

                //-------------------------------------------------------------------------------------------------------------------------------------
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Mes_Emp", OleDbType.VarChar).Value = txtMes_Empresa.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Anho_Emp", OleDbType.VarChar).Value = txtAnho_Empresa.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("NumFactProve", OleDbType.VarChar).Value = txtNroFactProvee.Text.ToUpper();

                int DataInsert = insertDataIntoMSAccessDataBaseOleDbCommand.ExecuteNonQuery();
                if (DataInsert > 0)
                {
                    //MessageBox.Show("Registro Exitosoo ¡¡¡");
                    MessageBox.Show("Registro Exitoso", "Registro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Refrescar el Datagrid al insertar un registro
                    ListaTablaGastos = PopulatedAddTablaGastos();
                    dataGridView1.DataSource = ListaTablaGastos;
                    //DarFormatoCabecera();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //cerramos la conexion
                CloseConexion();
            }
        }

        private void cbRUC_SelectedIndexChanged(object sender, EventArgs e)
        {
            string elruc = cbRUC.SelectedValue.ToString();

            List<TabClientes> ListaRUCResultado = new List<TabClientes>();

            ListaRUCResultado = ListaRUC.Where(x => x.ruc_cl == elruc).ToList();

            if (ListaRUCResultado.Count > 0)
            {
                txtNombEmpre.Text = ListaRUCResultado[0].nomb_emp.ToString();
            }
        }

        private void btnVIsualizar_Click(object sender, EventArgs e)
        {
            FormFact mv = new FormFact();
            mv.Show();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

            DateTime fecha1 = dateTimePicker1.Value;
            txtFechaRecibo.Text = fecha1.ToString();
            txtMes_Recibo.Text = fecha1.ToString("MM");
            txtAnho_Recibo.Text = fecha1.ToString("yyyy");
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha2 = dateTimePicker2.Value;
            txtFechaEmpresa.Text = fecha2.ToString();
            txtMes_Empresa.Text = fecha2.ToString("MM");
            txtAnho_Empresa.Text = fecha2.ToString("yyyy");
        }

        private void btnCleaner_Click(object sender, EventArgs e)
        {
            limpiar();
        }

        private void txtProyectoFiltr_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastosProyectFiltrar = new List<TablaGasto>();
            listaGastosProyectFiltrar = ListaTablaGastos.Where(item => item.proyecto.Contains(txtProyectoFiltr.Text)).ToList();
            dataGridView1.DataSource = listaGastosProyectFiltrar;
        }

        private void txtConceptFiltr_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoConceptFiltra = new List<TablaGasto>();
            listaGastoConceptFiltra = ListaTablaGastos.Where(item => item.concepto.Contains(txtConceptFiltr.Text)).ToList();
            dataGridView1.DataSource = listaGastoConceptFiltra;
        }

        private void txtPeriImpFiltr_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoPeriodoFiltr = new List<TablaGasto>();
            listaGastoPeriodoFiltr = ListaTablaGastos.Where(item => item.PeriImputaci.Contains(txtPeriImpFiltr.Text)).ToList();
            dataGridView1.DataSource = listaGastoPeriodoFiltr;
        }

        private void txtRUC_Client_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoRUC_Cliente = new List<TablaGasto>();
            listaGastoRUC_Cliente = ListaTablaGastos.Where(item => item.ruc_cl.Contains(txtRUC_Client_Filtro.Text)).ToList();
            dataGridView1.DataSource = listaGastoRUC_Cliente;
        }

        private void txtRuc_ProveeFiltr_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoRUC_Proveedor = new List<TablaGasto>();
            listaGastoRUC_Proveedor = ListaTablaGastos.Where(item => item.ruc_pr.Contains(txtRuc_ProveeFiltr.Text)).ToList();
            dataGridView1.DataSource = listaGastoRUC_Proveedor;
        }

        private void txtTipoFactFiltr_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoTipoFactura = new List<TablaGasto>();
            listaGastoTipoFactura = ListaTablaGastos.Where(item => item.tipo.Contains(txtTipoFactFiltr.Text)).ToList();
            dataGridView1.DataSource = listaGastoTipoFactura;
        }

        private void txtQuienPag_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoQuienPag = new List<TablaGasto>();
            listaGastoQuienPag = ListaTablaGastos.Where(item => item.quienPag.Contains(txtQuienPag_Filtro.Text)).ToList();
            dataGridView1.DataSource = listaGastoQuienPag;
        }

        private void txtMes_Factura_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listagastomesfacturafiltr = new List<TablaGasto>();
            listagastomesfacturafiltr = ListaTablaGastos.Where(item => item.Mes_Reci.Contains(txtMes_Factura_Filtro.Text)).ToList();
            dataGridView1.DataSource = listagastomesfacturafiltr;
        }

        private void txtMes_Empresa_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listaGastoMesPagoEmpresaFiltro = new List<TablaGasto>();
            listaGastoMesPagoEmpresaFiltro = ListaTablaGastos.Where(item => item.Mes_Emp.Contains(txtMes_Empresa_Filtro.Text)).ToList();
            dataGridView1.DataSource = listaGastoMesPagoEmpresaFiltro;
        }

        private void txtAnho_Empresa_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listagastoanhoempresafiltr = new List<TablaGasto>();
            listagastoanhoempresafiltr = ListaTablaGastos.Where(item => item.Anho_Emp.Contains(txtMes_Empresa_Filtro.Text)).ToList();
            dataGridView1.DataSource = listagastoanhoempresafiltr;
        }

        private void txtAnho_Factura_Filtro_TextChanged(object sender, EventArgs e)
        {
            List<TablaGasto> listagastomesfacturafiltr = new List<TablaGasto>();
            listagastomesfacturafiltr = ListaTablaGastos.Where(item => item.Anho_Reci.Contains(txtAnho_Factura_Filtro.Text)).ToList();
            dataGridView1.DataSource = listagastomesfacturafiltr;

        }
    }
}
