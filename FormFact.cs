using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using SpreadsheetLight;
using static facturacion.Entidades;
using Microsoft.VisualBasic;
using System.IO;

namespace facturacion
{
    public partial class FormFact : Form
    {
        // cadena de conexion string a la base de datos access
        String rutaAcces = Entidades.Variablesglobales.rutaAccess;
        String microsolftAccessDatabaseConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = ";  //LEER DE CUALQUIER DISCO C: D: F: etc
        //String microsolftAccessDatabaseConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = C:\Datos\fac_cliente.accdb";  //LEER DE CUALQUIER DISCO C: D: F: etc

        //Consulta a la bas de datos
        String selectDataFromMSAccessDatabaseQuery = "SELECT * FROM TablaFactura ";

        //query para insertar datos en base de datos access
        String InsertDataIntoMicrosoftAccessDatabase = "INSERT INTO TablaFactura (nomb_emp,Proyecto,tipofactura,numFactura,ruc_cl,import,IGV,conIgv,detraccion12,fechaCobroDetra,mes_CobDetra,anho_CobDetra,aCobrarEnCuenta,fechapagoCliente,mes_pagCli,anho_pagCli,entregaTrabajos,aprobacionTrabajos,fechaEmision,mes_Emi,anho_Emi,fechaVencimientoPago,fechaPrevista,observacionesPag, fechaInicioTrabajo,fechaInicioContrato,fechaFinalContrato,fechaRealTrabajo,fechaFinalPrevista) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        
        //TRAER EL VALOR DE RUC DE TABLA DE CLIENTES
        String QueryDatosTablaCliente = "SELECT * FROM TablaCliente";

        //Query para actualizar los datos de la base de datos de access
        String UpdateMicrosoftAccessDataQuery = "UPDATE TablaFactura SET nomb_emp = ? ,Proyecto = ? ,tipofactura = ?,numFactura = ?, ruc_cl = ?,import = ?,IGV = ?,conIgv = ?,detraccion12 = ?,fechaCobroDetra = ?,mes_CobDetra = ?,anho_CobDetra = ?,aCobrarEnCuenta = ?,fechapagoCliente = ?,mes_pagCli = ?,anho_pagCli = ?,entregaTrabajos = ?,aprobacionTrabajos = ?,fechaEmision = ?, mes_Emi = ?,anho_Emi = ?,fechaVencimientoPago =?,fechaPrevista= ?,observacionesPag = ?,fechaInicioTrabajo=?, fechaInicioContrato=?, fechaFinalContrato=? , fechaRealTrabajo=?,fechaFinalPrevista=? WHERE ID = ?";

        //Query para ELIMINAR los datos de la base de datos de access
        String DeleteMicrosoftAccessDataQuery = "DELETE * FROM TablaFactura WHERE ID = ?";

        //Se agrega una lista general
        List<TabFactura> listaFactura = new List<TabFactura>();
        List<TabClientes> ListaRUC = new List<TabClientes>();


        public OleDbConnection microsolftAccessDatabaseOleDbConnection = null;
        public FormFact()
        {
            //instanciamos la conexion a la base de datoa access
            // microsolftAccessDatabaseOleDbConnection = new OleDbConnection(microsolftAccessDatabaseConnectionString);
            //MessageBox.Show("Conectado base de datos Access ¡¡¡");
            InitializeComponent();
        }

        public void Openconectar()
        {
            //abrimos la conexión
           microsolftAccessDatabaseOleDbConnection = new OleDbConnection(microsolftAccessDatabaseConnectionString+rutaAcces);
            if (microsolftAccessDatabaseOleDbConnection.State == ConnectionState.Closed)
            {
                microsolftAccessDatabaseOleDbConnection.Open();
                //MessageBox.Show("Conexión Abierta ......");
            }

        }
        public void CloseDesconectar()
        {
            //cerramos la conexion
            if (microsolftAccessDatabaseOleDbConnection.State == ConnectionState.Open)
            {
                microsolftAccessDatabaseOleDbConnection.Close();
                //MessageBox.Show("Conexión Cerrada ......");
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
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

            xlWorkSheet.Columns[1].ColumnWidth = 5;
            xlWorkSheet.Columns[2].ColumnWidth = 60;
            xlWorkSheet.Columns[3].ColumnWidth = 60;
            xlWorkSheet.Columns[4].ColumnWidth = 20;


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

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha1 = dateTimePicker1.Value;
            txtFechaCobrodetra.Text = fecha1.ToString();
            txtMesDetra.Text = fecha1.ToString("MM");
            txtAnhoDetra.Text = fecha1.ToString("yyyy");
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha3 = dateTimePicker3.Value;
            txtEntregaTrabajos.Text = fecha3.ToString();
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha4 = dateTimePicker4.Value;
            txtAprobTrabajos.Text = fecha4.ToString();
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha5 = dateTimePicker5.Value;
            txtFechaEmision.Text = fecha5.ToString();
            txtMes_Emision.Text = fecha5.ToString("MM");
            txtAnho_Emision.Text = fecha5.ToString("yyyy");
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha6 = dateTimePicker6.Value;
            txtFechaVencimientoPago.Text = fecha6.ToString();
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha7 = dateTimePicker7.Value;
            txtFechaPrevista.Text = fecha7.ToString();
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            LimpiarCamposSolo();
            
            
            ///PARA BLOOQUEAR Y CAMBIAR DE NOMBRE EL BOTON
            //btnAgregar.Enabled = true;
            //btnAgregar.Text = "Copiar";
            
            
            
            txtID.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtNombEmp.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtProyecto.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtNumFactura.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            cbRucCliente.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            txtImporte.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            txtIGV_.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            txtConIgv.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            txtDetraccion12.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();

            //DateTime? fechacobroobtener = (DateTime) dataGridView1.CurrentRow.Cells[10].Value;
            //var fehcaCobroNulo = (object)DBNull.Value;

            if (dataGridView1.CurrentRow.Cells[10].Value != null) //valida en columna del datagridview  si es nullo que lo llene como STRING
            {
                DateTime dtFechaCobrodetra = (DateTime) dataGridView1.CurrentRow.Cells[10].Value;
                txtFechaCobrodetra.Text = dtFechaCobrodetra.ToString("dd/MM/yyyy");
            }

            txtMesDetra.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            txtAnhoDetra.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();
            txtCobrarenCuenta.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();

            if (dataGridView1.CurrentRow.Cells[14].Value != null) //valida en columna del datagridview  si es nullo que lo llene como STRING
            {
                DateTime dtFechaPagCliente = (DateTime)dataGridView1.CurrentRow.Cells[14].Value;
                txtFechaPagCliente.Text = dtFechaPagCliente.ToString("dd/MM/yyyy");
            }
            //----------------------------------------------------------------------------------------------------------------------------------------------
            txtMes_PagCliente.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            txtAnho_PagCliente.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
                
            //------------------------------------------------------------------------------------------------------------------------------------

            if (dataGridView1.CurrentRow.Cells[17].Value != null) //valida en columna del datagridview  si es nullo que lo llene como STRING
            {
                txtEntregaTrabajos.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            }
            //-------------------------------------------------------------------------------------------------------------------------------------

            if (dataGridView1.CurrentRow.Cells[18].Value != null) //valida en columna del datagridview  si es nullo que lo llene como STRING
            {
                txtAprobTrabajos.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[19].Value != null)
            {
                txtFechaEmision.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            }

            
           txtMes_Emision.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
           txtAnho_Emision.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
          

            //--------------------------------------------------------------------------------------------------------------------------------------

            if (dataGridView1.CurrentRow.Cells[22].Value !=null)
            {
                txtFechaVencimientoPago.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();
            }
            
            //--------------------------------------------------------------------------------------------------------------------------------------

            //-------------------------------------------------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[23].Value != null)
            {
                txtFechaPrevista.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            }
            
            //-------------------------------------------------------------------------------------------------------------------------------------
            txtObservacionesPago.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();

            //--------------------fechaInicioTrabajo---------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[25].Value != null)
            {
                txtInicTrabajo.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            }

            //--------------------fechaInicioContrato---------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[26].Value != null)
            {
                txtInicContrato.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
            }

            //--------------------fechaFinalContrato---------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[27].Value != null)
            {
                txtFinalContrato.Text = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            }

            //--------------------fechaRealTrabajo---------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[28].Value != null)
            {
                txtRealTrabajo.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();
            }

            //--------------------fechaFinalPrevista---------------------------------------------------------------------------------------------
            if (dataGridView1.CurrentRow.Cells[29].Value != null)
            {
                txtFinalPrevista.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
            }

            //--------------------------------------------------------------------------------------------------------------------------------

        }

        public List<TabFactura> populateDataGridViewFromMicrosoftAccessDatabase() //FUNCION PARA LEER LOS DATOS DE LA BASE DE DATOS
        {
         
            List<TabFactura> lista = new List<TabFactura>();
            int IDError = 0;
            int ColumnaError = 0;
            try
            {  
                if (dataGridView1.Rows.Count > 0)
                {
                   
                    dataGridView1.DataSource = null;
                }

                OleDbCommand populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(selectDataFromMSAccessDatabaseQuery, microsolftAccessDatabaseOleDbConnection);       
                Openconectar();      
                TabFactura objeto;

                OleDbDataReader reader = populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabFactura();
                    objeto.Id = (int)reader[0];
                    objeto.nomb_emp = reader[1].ToString();
                    objeto.proyecto = reader[2].ToString();                  
                    objeto.tipofactura = reader[3].ToString();
                
                    objeto.numFactura = reader[4].ToString();
                    ColumnaError = 6;
                    objeto.ruc_cl = reader[5].ToString();
                    ColumnaError = 7;
                    //Valores en Decimales
                    if (Information.IsDBNull(reader[6]) != true)
                    {
                        objeto.import = float.Parse(reader[6].ToString());
                    }

                    ColumnaError = 8;
                    if (Information.IsDBNull(reader[7]) != true)
                    {
                        objeto.IGV = float.Parse(reader[7].ToString());
                    }
                    ColumnaError = 9;
                    
                    if (Information.IsDBNull(reader[8]) != true)
                    {
                        objeto.conIgv = float.Parse(reader[8].ToString());
                    }
                    ColumnaError = 10;
                    
                    if (Information.IsDBNull(reader[9]) != true)
                    {
                        objeto.detraccion12 = float.Parse(reader[9].ToString());
                    }


                    //valores en fechas
                    // objeto.fechaCobroDetraConsulta = (DateTime)Interaction.IIf(Information.IsDBNull(reader[10]),DateTime.Now, reader[10]); //LEE LA FECHA ACTUAL DEL SISTEMA Y TE LO COLOCA EN DATAGRID
                    ColumnaError = 11;
                    if (Information.IsDBNull(reader[10]) != true)
                    {
                        objeto.fechaCobroDetra = (DateTime)reader[10];
                    }
                    ColumnaError = 12;
                    objeto.mes_CobDetra = reader[11].ToString();
                    ColumnaError = 13;
                    objeto.anho_CobDetra = reader[12].ToString();

                    ColumnaError = 14;
                    // -------------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[13]) != true)
                    {
                        objeto.aCobrarEnCuenta = float.Parse(reader[13].ToString());
                    }

                    ColumnaError = 15;
                    if (Information.IsDBNull(reader[14]) != true)
                    {
                        objeto.fechapagoCliente = (DateTime)reader[14];
                    }
                    ColumnaError = 16;
                    objeto.mes_pagCli = reader[15].ToString();
                    ColumnaError = 17;
                    objeto.anho_pagCli = reader[16].ToString();

                    ColumnaError = 18;
                    //-------------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[17]) != true)
                    {
                        objeto.entregaTrabajos = (DateTime)reader[17];
                    }
                    ColumnaError = 19;
                    //------------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[18]) != true)
                    {
                        objeto.aprobacionTrabajos = (DateTime)reader[18];
                    }
                    ColumnaError = 20;
                    //------------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[19]) != true)
                    {
                        objeto.fechaEmision = (DateTime)reader[19];
                    }
                    ColumnaError = 21;
                    objeto.mes_Emi = reader[20].ToString();
                    ColumnaError = 22;
                    objeto.anho_Emi = reader[21].ToString();

                    ColumnaError = 23;
                    //-----------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[22]) != true)
                    {
                        objeto.fechaVencimientoPago = (DateTime)reader[22];
                    }
                    //-----------------------------------------------------------------------------------------------------------------------------
                    ColumnaError = 24;
                    //-----------------------------------------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[23]) != true)
                    {
                        objeto.fechaPrevista = (DateTime)reader[23];
                    }
                    ColumnaError = 25;
                    
                    //OBSERVACIONES
                    objeto.observacionesPag = reader[24].ToString();
                    ColumnaError = 26;

                    //--------------------fechaInicioTrabajo-------------------------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[25]) != true)
                    {
                        objeto.fechaInicioTrabajo = (DateTime)reader[25];
                    }
                    ColumnaError = 27;

                    //--------------------fechaInicioContrato------------------------------------------------------
                    if (Information.IsDBNull(reader[26]) != true)
                    {
                        objeto.fechaInicioContrato = (DateTime)reader[26];
                    }
                    ColumnaError = 28;

                    //----------------------fechaFinalContrato---------------------------------------------------------------------------
                    if (Information.IsDBNull(reader[27]) != true)
                    {
                        objeto.fechaFinalContrato = (DateTime)reader[27];
                    }
                    ColumnaError = 29;

                    //------------------------fechaRealTrabajo--------------------------------------------------------

                    if (Information.IsDBNull(reader[28]) != true)
                    {
                        objeto.fechaRealTrabajo = (DateTime)reader[28];
                    }
                    ColumnaError = 30;

                    //--------------------fechaFinalPrevista--------------------------------------------------------------------


                    if (Information.IsDBNull(reader[29]) != true)
                    {
                        objeto.fechaFinalPrevista = (DateTime)reader[29];
                    }
                    ColumnaError = 31;

                    lista.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                string mensajeErrror = "El codigo con : " + IDError.ToString() + " Tiene un error. La columna es: " +  ColumnaError.ToString();

                MessageBox.Show(ex.Message+ " " +mensajeErrror);
                return lista;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseDesconectar();
            }
            return lista;
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


        public List<TabClientes> PopulatedComboRUC()
        {
            //control k + control s  : sacar el TRY CATH
            //creamos la lista en una instania de lista de la tabla Factura
            List<TabClientes> listaRUC = new List<TabClientes>();
            try
            {
                //// borrar filas de cuadrícula de datos antes de cargar datos de accesos de microsoft
                if (cbRucCliente.Items.Count > 0)
                {
                    cbRucCliente.DataSource = null;
                }
                OleDbCommand conceptoMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryDatosTablaCliente, microsolftAccessDatabaseOleDbConnection);
                //abrimos la conexion a base de datos
                Openconectar();

                //Es el nombre de la variable , se crea para poder instanciar y poder llenar los datos y
                //para que se guarden de forma ordenada se crea un lista
                TabClientes objeto;

                OleDbDataReader reader = conceptoMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                // bucle para leer los datos de Microsoft Access Database
                while (reader.Read())
                {
                    objeto = new TabClientes();

                    objeto.ruc_cl = reader[1].ToString();
                    objeto.nomb_emp = reader[2].ToString();

                    listaRUC.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return listaRUC;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseDesconectar();
            }
            return listaRUC;
        }


        public void llenarRUC()
        {
            Openconectar();
            ListaRUC = PopulatedComboRUC();
            cbRucCliente.DisplayMember = "ruc_cl";
            cbRucCliente.ValueMember = "ruc_cl";

        }
        public void IniciarDatos()
        {
            Openconectar();
            //dataGridView1.RowTemplate.Height = 30;
            //dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightBlue;
            //populateDataGridViewFromMicrosoftAccessDatabase();
            listaFactura = populateDataGridViewFromMicrosoftAccessDatabase();
            llenarRUC();
            dataGridView1.DataSource = listaFactura;

            TabClientes objectoSele = new TabClientes();
            objectoSele.ruc_cl = "Selecionar";
            objectoSele.nomb_emp = "";
            ListaRUC.Add(objectoSele);
            int indexValores = ListaRUC.Count;
            cbRucCliente.DataSource = ListaRUC;
            cbRucCliente.SelectedIndex = indexValores - 1;

            DarFormatoCabecera();
            cargarDatosCombobox();
        }

        public void DarFormatoCabecera()
        {
            DataGridViewColumn columna0 = dataGridView1.Columns[0];
            columna0.HeaderText = "ID";
            columna0.Width = 30;

            DataGridViewColumn columna1 = dataGridView1.Columns[1];
            columna1.HeaderText = "Nombre de la Empresa";
            columna1.Width = 300;

            DataGridViewColumn columna2 = dataGridView1.Columns[2];
            columna2.HeaderText = "Proyecto";
            columna2.Width = 300;

            DataGridViewColumn columna3 = dataGridView1.Columns[3];
            columna3.HeaderText = "Tipo de Factura";
            columna3.Width = 80;

            DataGridViewColumn columna4 = dataGridView1.Columns[4];
            columna4.HeaderText = "Numero de Factura";
            columna4.Width = 100;

            DataGridViewColumn columna5 = dataGridView1.Columns[5];
            columna5.HeaderText = "RUC-Cliente";
            columna5.Width = 100;

            DataGridViewColumn columna6 = dataGridView1.Columns[6];
            columna6.HeaderText = "Importe";
            columna6.Width = 50;

            DataGridViewColumn columna7 = dataGridView1.Columns[7];
            columna7.HeaderText = "IGV";
            columna7.Width = 50;

            DataGridViewColumn columna8 = dataGridView1.Columns[8];
            columna8.HeaderText = "Con IGV";
            columna8.Width = 50;

            DataGridViewColumn columna9 = dataGridView1.Columns[9];
            columna9.HeaderText = "Detraccion 12%";
            columna9.Width = 70;
            //-------------------------------- FECHA DE COBRO DE DETRACCIONES ---------------------------------
            DataGridViewColumn columna10 = dataGridView1.Columns[10];
            columna10.HeaderText = "Fecha de Cobro de las Detracciones";
            columna10.Width = 150;

            DataGridViewColumn columna11 = dataGridView1.Columns[11];
            columna11.HeaderText = "Mes";
            columna11.Width = 50;
        
            DataGridViewColumn columna12 = dataGridView1.Columns[12];
            columna12.HeaderText = "Año";
            columna12.Width = 80;
            //--------------------------------------------------------------------------------------------------------
            DataGridViewColumn columna13 = dataGridView1.Columns[13];
            columna13.HeaderText = "Cobró Neto - Cuenta (Pagó Final C/desc.SUNAT)";
            columna13.Width = 120;

            //----------------------------- FECHA DE PAGO CLIENTE -----------------------------------

            DataGridViewColumn columna14 = dataGridView1.Columns[14];
            columna14.HeaderText = "Fecha de Pago Cliente";
            columna14.Width = 150;

            DataGridViewColumn columna15 = dataGridView1.Columns[15];
            columna15.HeaderText = "Mes";
            columna15.Width = 50;

            DataGridViewColumn columna16 = dataGridView1.Columns[16];
            columna16.HeaderText = "Año";
            columna16.Width = 80;
            //-------------------------------------------------------------------------------------------

            DataGridViewColumn columna17 = dataGridView1.Columns[17];
            columna17.HeaderText = "Fecha de Entrega de los Trabajos";
            columna17.Width = 100;

            DataGridViewColumn columna18 = dataGridView1.Columns[18];
            columna18.HeaderText = "Fecha de Aprobación de los Trabajos";
            columna18.Width = 100;

            //--------------------------FECHA DE EMISIÓN --------------------------------------------
            DataGridViewColumn columna19 = dataGridView1.Columns[19];
            columna19.HeaderText = "Fecha de Emisión";
            columna19.Width = 150;

            DataGridViewColumn columna20 = dataGridView1.Columns[20];
            columna20.HeaderText = "Mes";
            columna20.Width = 50;

            DataGridViewColumn columna21 = dataGridView1.Columns[21];
            columna21.HeaderText = "Año";
            columna21.Width = 80;

            //--------------------------------------------------------------------------------------

            DataGridViewColumn columna22 = dataGridView1.Columns[22];
            columna22.HeaderText = "Fecha de Vencimiento de Pagó";
            columna22.Width = 100;

           
            DataGridViewColumn columna23 = dataGridView1.Columns[23];
            columna23.HeaderText = "Fecha Prevista";
            columna23.Width = 100;

            DataGridViewColumn columna24 = dataGridView1.Columns[24];
            columna24.HeaderText = "Observaciones de Pagó";
            columna24.Width = 200;

            //-----------------------------------------------------------------------------------------------------------------------

            //DataGridViewColumn columna25 = dataGridView1.Columns[25];
            //columna25.HeaderText = "Fecha Inicio Trabajo";
            //columna25.Width = 100;

            //DataGridViewColumn columna26 = dataGridView1.Columns[26];
            //columna26.HeaderText = "Fecha Inicio Contrato";
            //columna26.Width = 100;


            //DataGridViewColumn columna27 = dataGridView1.Columns[25];
            //columna27.HeaderText = "Fecha Final Contrato";
            //columna27.Width = 100;


            //DataGridViewColumn columna28 = dataGridView1.Columns[28];
            //columna28.HeaderText = "Fecha Real Trabajo";
            //columna28.Width = 100;


            //DataGridViewColumn columna29 = dataGridView1.Columns[29];
            //columna29.HeaderText = "Fecha Final Trabajo";
            //columna29.Width = 100;

        }
      
        private void btnDeconectDDBB_Click(object sender, EventArgs e)
        {
            CloseDesconectar();
            LimpiarCampos();
            lblRuta.Text = "";
            MessageBox.Show("Desconectado de la Base de datos .......");
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
            //Abrir conexion
            Openconectar();
            OleDbCommand insertDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(InsertDataIntoMicrosoftAccessDatabase, microsolftAccessDatabaseOleDbConnection);

            //condicional para que evalue si todos estan llenos
            if (txtNombEmp.Text == String.Empty || txtProyecto.Text == String.Empty || comboBox1.Text == String.Empty || txtNumFactura.Text == String.Empty || cbRucCliente.Text == String.Empty || txtImporte.Text == String.Empty || txtIGV_.Text == String.Empty || txtConIgv.Text == String.Empty ||
                txtDetraccion12.Text == String.Empty || txtCobrarenCuenta.Text == String.Empty || txtFechaPagCliente.Text == String.Empty || txtEntregaTrabajos.Text == String.Empty || txtAprobTrabajos.Text == String.Empty || txtFechaEmision.Text == String.Empty ||
                txtFechaVencimientoPago.Text == String.Empty || txtMes_Emision.Text == String.Empty || txtAnho_Emision.Text == String.Empty || txtFechaPrevista.Text == String.Empty || txtObservacionesPago.Text == String.Empty)

            {
                //MessageBox.Show("Verificar que uno o más campos vacíos estén llenos............");

            }
            try
            {
                //los nombres de los campos debe ser correspondientes con la base de datos Access
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmp.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtProyecto.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("tipofactura", OleDbType.VarChar).Value = comboBox1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("numFactura", OleDbType.VarChar).Value = txtNumFactura.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = cbRucCliente.Text;

                //VALOR DEL IMPORTE
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import", OleDbType.VarChar).Value = txtImporte.Text;

                //CALCULO DEL IGV
                double calculoIgv;
                calculoIgv = (Convert.ToDouble(txtImporte.Text)) * 0.18;
                txtIGV_.Text = System.Convert.ToString(calculoIgv);
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("IGV", OleDbType.VarChar).Value = txtIGV_.Text;

                //CALCULO DEL IMPORTE+IGV = ConIgv
                double calculoConIgv;
                calculoConIgv = (Convert.ToDouble(txtImporte.Text)) + (Convert.ToDouble(txtIGV_.Text));
                txtConIgv.Text = System.Convert.ToString(calculoConIgv);
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("conIgv", OleDbType.VarChar).Value = txtConIgv.Text;


                //CALCULO DE DETRACCION
                double CalculoDetraccion12;
                CalculoDetraccion12 = (Convert.ToDouble(txtConIgv.Text)) * 0.12;
                txtDetraccion12.Text = System.Convert.ToString(CalculoDetraccion12);
                if ((Convert.ToDouble(txtImporte.Text)) > 700)
                {
                    insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("detraccion12", OleDbType.VarChar).Value = txtDetraccion12.Text;
                }

                //lectura de valores nulos
                DateTime fechaCobroDetra;
                var fechaCobrodetraNulo = (object)DBNull.Value;
                if (txtFechaCobrodetra.Text != "")
                {
                    if (DateTime.TryParse(txtFechaCobrodetra.Text, out fechaCobroDetra))
                    {
                        fechaCobrodetraNulo = fechaCobroDetra.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaCobroDetra", OleDbType.VarChar).Value = fechaCobrodetraNulo;
                //---------------------------------------------------------------------------------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_CobDetra", OleDbType.VarChar).Value = txtMes_Emision.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_CobDetra", OleDbType.VarChar).Value = txtAnho_Emision.Text;


                //CALCULO A COBRAR EN CUENTA
                double CalculoCobrarenCuenta;
                CalculoCobrarenCuenta = (Convert.ToDouble(txtConIgv.Text)) - (Convert.ToDouble(txtDetraccion12.Text));
                txtCobrarenCuenta.Text = System.Convert.ToString(CalculoCobrarenCuenta);
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("aCobrarEnCuenta", OleDbType.VarChar).Value = txtCobrarenCuenta.Text;


                //FECHA DE COBRO
                DateTime fechaCobro;
                var fechacobronull = (object)DBNull.Value;
                if (txtFechaPagCliente.Text != "")
                {
                    if (DateTime.TryParse(txtFechaPagCliente.Text, out fechaCobro))
                    {
                        fechacobronull = fechaCobro.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechapagoCliente", OleDbType.VarChar).Value = fechacobronull;

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_pagCli", OleDbType.VarChar).Value = txtMes_PagCliente.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_pagCli", OleDbType.VarChar).Value = txtAnho_PagCliente.Text;


                //FECHA DE ENTREGA DE TRABAJOS
                DateTime fechaEntregaTrabajos;
                var fechEntregTrabajNull = (object)DBNull.Value;
                if (txtEntregaTrabajos.Text != "")
                {
                    if (DateTime.TryParse(txtEntregaTrabajos.Text, out fechaEntregaTrabajos))
                    {
                        fechEntregTrabajNull = fechaEntregaTrabajos.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("entregaTrabajos", OleDbType.VarChar).Value = fechEntregTrabajNull;

                //FECHA DE APROBACION DE TRABAJOS
                DateTime fechaAprobTrabajos;
                var fechaAprobTrabjNull = (object)DBNull.Value;
                if (txtAprobTrabajos.Text != "")
                {
                    if (DateTime.TryParse(txtAprobTrabajos.Text, out fechaAprobTrabajos))
                    {
                        fechaAprobTrabjNull = fechaAprobTrabajos.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("aprobacionTrabajos", OleDbType.VarChar).Value = fechaAprobTrabjNull;

                //FECHA DE EMISION
                DateTime fechaEmision;
                var fechaEmisionNull = (object)DBNull.Value;
                if (txtFechaEmision.Text != "")
                {
                    if (DateTime.TryParse(txtFechaEmision.Text, out fechaEmision))
                    {
                        fechaEmisionNull = fechaEmision.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaEmision", OleDbType.VarChar).Value = fechaEmisionNull;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_Emi", OleDbType.VarChar).Value = txtMes_Emision.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_Emi", OleDbType.VarChar).Value = txtAnho_Emision.Text;

                //FECHA DE VENCIMIENTO DE PAGO
                DateTime fechaVencimientPago;
                var fechaVenciPagoNull = (object)DBNull.Value;
                if (txtFechaVencimientoPago.Text != null)
                {
                    if (DateTime.TryParse(txtFechaVencimientoPago.Text, out fechaVencimientPago))
                    {
                        fechaVenciPagoNull = fechaVencimientPago.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaVencimientoPago", OleDbType.VarChar).Value = fechaVenciPagoNull;


                //FECHA PREVISTA DE PAGO
                DateTime fechaPrevista;
                var fechaPrevistaNull = (object)DBNull.Value;
                if (txtFechaPrevista.Text != null)
                {
                    if (DateTime.TryParse(txtFechaPrevista.Text, out fechaPrevista))
                    {
                        fechaPrevistaNull = fechaPrevista.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaPrevista", OleDbType.VarChar).Value = fechaPrevistaNull;



                //OBSERVACIONES DE DE PAGO
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("observacionesPag", OleDbType.VarChar).Value = txtObservacionesPago.Text;


                //___________________________________________________________________________________________________________________________

                //FECHA INICIO TRABAJO
                DateTime fechaInicioTrabj;
                var fechaInicioTrabajNull = (object)DBNull.Value;
                if (txtInicTrabajo.Text != null)
                {
                    if (DateTime.TryParse(txtInicTrabajo.Text, out fechaInicioTrabj))
                    {
                        fechaInicioTrabajNull = fechaInicioTrabj.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaInicioTrabajo", OleDbType.VarChar).Value = fechaInicioTrabajNull;


                //FECHA INICIO CONTRATO
                DateTime fechaInicioContrat;
                var fechaInicioContratNull = (object)DBNull.Value;
                if (txtInicContrato.Text != null)
                {
                    if (DateTime.TryParse(txtInicContrato.Text, out fechaInicioContrat))
                    {
                        fechaInicioContratNull = fechaInicioContrat.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaInicioContrato", OleDbType.VarChar).Value = fechaInicioContratNull;


                //FECHA FINAL CONTRATO
                DateTime fechaFinalContrato;
                var fechaFinalContratoNull = (object)DBNull.Value;
                if (txtFinalContrato.Text != null)
                {
                    if (DateTime.TryParse(txtFinalContrato.Text, out fechaFinalContrato))
                    {
                        fechaFinalContratoNull = fechaFinalContrato.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaFinalContratoNull;


                //FECHA REAL CONTRATO
                DateTime fechaRealTrabaj;
                var fechaRealTrabajNull = (object)DBNull.Value;
                if (txtRealTrabajo.Text != null)
                {
                    if (DateTime.TryParse(txtRealTrabajo.Text, out fechaRealTrabaj))
                    {
                        fechaRealTrabajNull = fechaRealTrabaj.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaRealTrabajNull;


                //FECHA FINAL PREVISTA
                DateTime fechaFinalPrevist;
                var fechaFinalPrevistNull = (object)DBNull.Value;
                if (txtFechaPrevista.Text != null)
                {
                    if (DateTime.TryParse(txtFechaPrevista.Text, out fechaFinalPrevist))
                    {
                        fechaFinalPrevistNull = fechaFinalPrevist.ToString("dd/MM/yyyy");
                    }
                }
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaFinalPrevistNull;

                //---------------------------------------------------------------------------------------
                
                int DataInsert = insertDataIntoMSAccessDataBaseOleDbCommand.ExecuteNonQuery();
                if (DataInsert > 0)
                {
                    //MessageBox.Show("Registro Exitosoo ¡¡¡");
                    MessageBox.Show("Registro Exitoso", "Registro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Refrescar el Datagrid al insertar un registro
                    listaFactura = populateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = listaFactura;
                    DarFormatoCabecera();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //cerramos la conexion
                CloseDesconectar();
            }
        }



        //ACTUALIZACION DE DATAGRIDVIEW
        private void btnActualizar_Click(object sender, EventArgs e)
        {
            
            if (String.IsNullOrEmpty(txtID.Text))
            {
                MessageBox.Show("Hacer click en uno de los registros , que desee actualizar");
            }
            else
            {
                //abrir conexion
                Openconectar();
                OleDbCommand UpdateDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(UpdateMicrosoftAccessDataQuery, microsolftAccessDatabaseOleDbConnection); 
                //condicional para que evalue si todos estan llenos

                try
                {
                    //los nombres de los campos debe ser correspondientes con la base de datos Access

                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmp.Text;             
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtProyecto.Text; 
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("tipofactura", OleDbType.VarChar).Value = comboBox1.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("numFactura", OleDbType.VarChar).Value = txtNumFactura.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = cbRucCliente.Text;
                    

                    //VALOR DEL IMPORTE
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import", OleDbType.VarChar).Value = txtImporte.Text;
                   
                    //CALCULO DEL IGV
                    double calculoIgv;
                    calculoIgv = (Convert.ToDouble(txtImporte.Text)) * 0.18;
                    txtIGV_.Text = System.Convert.ToString(calculoIgv);
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("IGV", OleDbType.VarChar).Value = txtIGV_.Text;
                    

                    //CALCULO DEL IMPORTE+IGV = ConIgv
                    double calculoConIgv;
                    calculoConIgv = (Convert.ToDouble(txtImporte.Text)) + (Convert.ToDouble(txtIGV_.Text));
                    txtConIgv.Text = System.Convert.ToString(calculoConIgv);
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("conIgv", OleDbType.VarChar).Value = txtConIgv.Text;
                    

                    //CALCULO DE DETRACCION
                    double CalculoDetraccion12;
                    CalculoDetraccion12 = (Convert.ToDouble(txtConIgv.Text)) * 0.12;
                    txtDetraccion12.Text = System.Convert.ToString(CalculoDetraccion12);
                    if ((Convert.ToDouble(txtImporte.Text)) > 700)
                    {
                        UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("detraccion12", OleDbType.VarChar).Value = txtDetraccion12.Text;
                      
                    }

                    //-----------------------------------------------------------------------------------------------------------
                    DateTime fechaCobroDetra;
                    var fechaCobrodetraNulo = (object)DBNull.Value;
                    if (txtFechaCobrodetra.Text != "")
                    {
                        if (DateTime.TryParse(txtFechaCobrodetra.Text, out fechaCobroDetra))
                        {
                            fechaCobrodetraNulo = fechaCobroDetra.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaCobroDetra", OleDbType.VarChar).Value = fechaCobrodetraNulo;
                    
                    //---------------------------------------------------------------------------------------------------------------------------------------------
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_CobDetra", OleDbType.VarChar).Value = txtMes_Emision.Text;
                    
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_CobDetra", OleDbType.VarChar).Value = txtAnho_Emision.Text;
                    

                    //CALCULO A COBRAR EN CUENTA
                    double CalculoCobrarenCuenta;
                    CalculoCobrarenCuenta = (Convert.ToDouble(txtConIgv.Text)) - (Convert.ToDouble(txtDetraccion12.Text));
                    txtCobrarenCuenta.Text = System.Convert.ToString(CalculoCobrarenCuenta);
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("aCobrarEnCuenta", OleDbType.VarChar).Value = txtCobrarenCuenta.Text;
                    

                    //FECHA DE COBRO - ACTUALIZACIONES
                    DateTime fechaCobro;
                    var fechaCobroNulo = (object)DBNull.Value;
                    if (txtFechaPagCliente.Text != "")
                    {
                        if (DateTime.TryParse(txtFechaPagCliente.Text, out fechaCobro))
                        {
                            fechaCobroNulo = fechaCobro.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechapagoCliente", OleDbType.VarChar).Value = fechaCobroNulo;
                  
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_pagCli", OleDbType.VarChar).Value = txtMes_PagCliente.Text;
                  
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_pagCli", OleDbType.VarChar).Value = txtAnho_PagCliente.Text;
                    
                    //---------------------------------------------------------------------------------------------------------------------------------------------

                    //FECHA DE ENTREGA DE TRABAJOS //--------------------------------------------------------------------------------------------------------------
                    DateTime fechaEntregaTrabajos;
                    var fechEntregTrabajNull = (object)DBNull.Value;
                    if (txtEntregaTrabajos.Text != "")
                    {
                        if (DateTime.TryParse(txtEntregaTrabajos.Text, out fechaEntregaTrabajos))
                        {
                            fechEntregTrabajNull = fechaEntregaTrabajos.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("entregaTrabajos", OleDbType.VarChar).Value = fechEntregTrabajNull;
                    
                    //---------------------------------------------------------------------------------------------------------------------------------------

                    //FECHA DE ENTREGA DE TRABAJOS-------------------------------------------------------------------------------------------------------------
                    DateTime fechaAprobTrabajos;
                    var fechaAprobTrabjNull = (object)DBNull.Value;
                    if (txtAprobTrabajos.Text != "")
                    {
                        if (DateTime.TryParse(txtAprobTrabajos.Text, out fechaAprobTrabajos))
                        {
                            fechaAprobTrabjNull = fechaAprobTrabajos.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("aprobacionTrabajos", OleDbType.VarChar).Value = fechaAprobTrabjNull;
                    

                    //------------------------------------------------------------------------------------------------------------------------------------------------

                    //FECHA DE EMISION ACTUALIZACION
                    DateTime fechaEmi;
                    var fechaEmiNull = (object)DBNull.Value;
                    if (txtFechaEmision.Text != "")
                    {
                        if (DateTime.TryParse(txtFechaEmision.Text,out fechaEmi))
                        {
                            fechaEmiNull = fechaEmi.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaEmision", OleDbType.VarChar).Value = fechaEmiNull;
                    
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes_Emi", OleDbType.VarChar).Value = txtMes_Emision.Text;
                   
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho_Emi", OleDbType.VarChar).Value = txtAnho_Emision.Text;
                    

                    //-------------------------------------------------------------------------------------------------------------------------------------------------

                    //FECHA DE VENCIMIENTO PAGO - ACTUALIZACION

                    DateTime fechaVenciPago;
                    var fechaVenciPagoNull = (object)DBNull.Value;
                    if (txtFechaVencimientoPago.Text !="")
                    {
                        if (DateTime.TryParse(txtFechaVencimientoPago.Text,out fechaVenciPago))
                        {
                            fechaVenciPagoNull = fechaVenciPago.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaVencimientoPago", OleDbType.VarChar).Value = fechaVenciPagoNull;
                    
                    //---------------------------------------------------------------------------------------------------------------------
                    DateTime fechaPrevi;
                    var fechaPreviNull = (object)DBNull.Value;
                    if (txtFechaPrevista.Text != "")
                    {
                        if (DateTime.TryParse(txtFechaPrevista.Text,out fechaPrevi))
                        {
                            fechaPreviNull = fechaPrevi.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaPrevista", OleDbType.VarChar).Value = fechaPreviNull;
                    
                    //--------------------------------------------------------------------------------------------------------------------------------------------

                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("observacionesPag", OleDbType.VarChar).Value = txtObservacionesPago.Text;
                   

                    //FECHA INICIO TRABAJO
                    DateTime fechaInicioTrabj;
                    var fechaInicioTrabajNull = (object)DBNull.Value;
                    if (txtInicTrabajo.Text != null)
                    {
                        if (DateTime.TryParse(txtInicTrabajo.Text, out fechaInicioTrabj))
                        {
                            fechaInicioTrabajNull = fechaInicioTrabj.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaInicioTrabajo", OleDbType.VarChar).Value = fechaInicioTrabajNull;
                   

                    //FECHA INICIO CONTRATO
                    DateTime fechaInicioContrat;
                    var fechaInicioContratNull = (object)DBNull.Value;
                    if (txtInicContrato.Text != null)
                    {
                        if (DateTime.TryParse(txtInicContrato.Text, out fechaInicioContrat))
                        {
                            fechaInicioContratNull = fechaInicioContrat.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaInicioContrato", OleDbType.VarChar).Value = fechaInicioContratNull;
                    

                    //FECHA FINAL CONTRATO
                    DateTime fechaFinalContrato;
                    var fechaFinalContratoNull = (object)DBNull.Value;
                    if (txtFinalContrato.Text != null)
                    {
                        if (DateTime.TryParse(txtFinalContrato.Text, out fechaFinalContrato))
                        {
                            fechaFinalContratoNull = fechaFinalContrato.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaFinalContratoNull;
                    

                    //FECHA REAL CONTRATO
                    DateTime fechaRealTrabaj;
                    var fechaRealTrabajNull = (object)DBNull.Value;
                    if (txtRealTrabajo.Text != null)
                    {
                        if (DateTime.TryParse(txtRealTrabajo.Text, out fechaRealTrabaj))
                        {
                            fechaRealTrabajNull = fechaRealTrabaj.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaRealTrabajNull;
                    

                    //FECHA FINAL PREVISTA
                    DateTime fechaFinalPrevist;
                    var fechaFinalPrevistNull = (object)DBNull.Value;
                    if (txtFechaPrevista.Text != null)
                    {
                        if (DateTime.TryParse(txtFechaPrevista.Text, out fechaFinalPrevist))
                        {
                            fechaFinalPrevistNull = fechaFinalPrevist.ToString("dd/MM/yyyy");
                        }
                    }
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fechaFinalContrato", OleDbType.VarChar).Value = fechaFinalPrevistNull;
                    
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ID", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);
                    

                    int DataUpdate = UpdateDataIntoMSAccessDataBaseOleDbCommand.ExecuteNonQuery();
                    if (DataUpdate > 0)
                    {
                        MessageBox.Show("Actualizado Exitosoo ¡¡¡");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    
                }
                finally
                {
                                       
                    //Se debe refrescar el datagrid antes de actualizar los datos de la base de datos de ACCESSS
                    listaFactura = populateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = listaFactura;
                    DarFormatoCabecera();

                    CloseDesconectar();
                }


            }

        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            //primero VERIFICAR que si ID esta vacio (caja texto)
            Openconectar();

            try
            {
                if (String.IsNullOrEmpty(txtID.Text))
                {
                    MessageBox.Show("Campo ID Vacio , seleccionar un fila de registro");
                }
                else
                {
                    OleDbCommand DeleteMicrosoftAccessDataQueryOleDbCommand = new OleDbCommand(DeleteMicrosoftAccessDataQuery, microsolftAccessDatabaseOleDbConnection);
                    DeleteMicrosoftAccessDataQueryOleDbCommand.Parameters.AddWithValue("ID", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);
                    //abrir la coneexion


                    int DeleteMicrosoftAccessData = DeleteMicrosoftAccessDataQueryOleDbCommand.ExecuteNonQuery();
                    if (DeleteMicrosoftAccessData > 0)
                    {
                        MessageBox.Show("Registro Borrado");
                        //Se debe refrescar el datagrid antes de actualizar los datos eliminados de la base de datos de ACCESSS
                        //limpiar los capos
                        LimpiarCampos();
                        listaFactura = populateDataGridViewFromMicrosoftAccessDatabase();
                        dataGridView1.DataSource = listaFactura;
                        DarFormatoCabecera();
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {
                //Cerrar la conexin a la base de datos
                CloseDesconectar();
            }

        }


        private void cargarDatosCombobox()
        {
            comboBox1.Items.Add("FACTURA");
            comboBox1.Items.Add("PREVISION");
           
        }

        private void btnBuscar_Click(object sender, EventArgs e)
        {
            string searchValue = txtNombEmp.Text;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[1].Value.ToString().Equals(searchValue))
                    {
                        row.Selected = true;
                        break;
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void LimpiarCamposSolo()
        {

            txtID.Text = string.Empty;
            txtNombEmp.Text = string.Empty;
            txtProyecto.Text = string.Empty;
            comboBox1.Text = string.Empty;
            txtNumFactura.Text = string.Empty;
            cbRucCliente.Text = string.Empty;
            txtImporte.Text = string.Empty;
            txtIGV_.Text = string.Empty;
            txtConIgv.Text = string.Empty;
            txtDetraccion12.Text = string.Empty;
            txtFechaCobrodetra.Text = string.Empty;
            txtCobrarenCuenta.Text = string.Empty;
            txtFechaPagCliente.Text = string.Empty;
            txtEntregaTrabajos.Text = string.Empty;
            txtAprobTrabajos.Text = string.Empty;
            txtFechaEmision.Text = string.Empty;
            txtFechaVencimientoPago.Text = string.Empty;
            txtMes_Emision.Text = string.Empty;
            txtAnho_Emision.Text = string.Empty;
            txtFechaPrevista.Text = string.Empty;
            txtObservacionesPago.Text = string.Empty;
            txtInicTrabajo.Text = string.Empty;
            txtInicContrato.Text = string.Empty;
            txtFinalContrato.Text = string.Empty;
            txtRealTrabajo.Text = string.Empty;
            txtFinalPrevista.Text = string.Empty;

        }

        private void LimpiarCampos()
        {
            LimpiarCamposSolo();
            //dataGridView1.DataSource = null;
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
        }
 
        private void txtNombEmp1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroNombre = new List<TabFactura>();
            listaFacturaFiltroNombre = listaFactura.Where(item => item.nomb_emp.Contains(txtNombEmp1.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltroNombre;
        }

        private void txtProyecto1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroProyecto = new List<TabFactura>();
            listaFacturaFiltroProyecto = listaFactura.Where(item => item.proyecto.Contains(txtProyecto1.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listaFacturaFiltroProyecto;
        }

        private void txtMes1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroMes = new List<TabFactura>();
            listaFacturaFiltroMes = listaFactura.Where(item => item.mes_Emi.Contains(txtMes_Emi1.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listaFacturaFiltroMes;
        }

        private void txtAnho1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroAnhoEmi = new List<TabFactura>();
            listaFacturaFiltroAnhoEmi = listaFactura.Where(item => item.anho_Emi.Contains(txtAnho_Emi1.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listaFacturaFiltroAnhoEmi;
        }

        
        private void txtImporte_TextChanged(object sender, EventArgs e)
        {
            txtIGV_.Text = string.Empty;
            txtConIgv.Text = string.Empty;
            txtDetraccion12.Text = string.Empty;
            txtCobrarenCuenta.Text = string.Empty;
        }

        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha6 = dateTimePicker9.Value;
            txtFechaPagCliente.Text = fecha6.ToString();
            txtMes_PagCliente.Text = fecha6.ToString("MM");
            txtAnho_PagCliente.Text = fecha6.ToString("yyyy");
        }

     

        private void txtPagFactMes1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltrFactMes = new List<TabFactura>();
            listaFacturaFiltrFactMes = listaFactura.Where(item => item.mes_pagCli.Contains(txtPagFactMes1.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltrFactMes;
        }

        private void txtPagFactAnho1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltrFactAnho = new List<TabFactura>();
            listaFacturaFiltrFactAnho = listaFactura.Where(item => item.anho_pagCli.Contains(txtPagFactAnho1.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltrFactAnho;
        }

        private void txtPagDetraMes1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltrPagDetraMes = new List<TabFactura>();
            listaFacturaFiltrPagDetraMes = listaFactura.Where(item => item.mes_CobDetra.Contains(txtPagDetraMes1.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltrPagDetraMes;
        }

        private void txtPagDetraAnho1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> FiltroListaDetraAnho = new List<TabFactura>();
            FiltroListaDetraAnho = listaFactura.Where(item => item.anho_CobDetra.Contains(txtPagDetraAnho1.Text)).ToList();
            dataGridView1.DataSource = FiltroListaDetraAnho;
        }

        private void cbRucCliente_SelectedIndexChanged(object sender, EventArgs e)
        {
            string elruc = cbRucCliente.SelectedValue.ToString();


            List<TabClientes> ListaRUCResultado = new List<TabClientes>();

            ListaRUCResultado = ListaRUC.Where(x => x.ruc_cl == elruc).ToList();

            if (ListaRUCResultado.Count > 0)
            {
                txtNombEmp.Text = ListaRUCResultado[0].nomb_emp.ToString();
            }

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            DateTime fechaInicio = dateTimePicker2.Value;
            txtInicTrabajo.Text = fechaInicio.ToString();
           
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            DateTime fechaIncioContrato = dateTimePicker8.Value;
            txtInicContrato.Text = fechaIncioContrato.ToString();
        }

        private void dateTimePicker10_ValueChanged(object sender, EventArgs e)
        {
            DateTime fechaFinalContrato = dateTimePicker10.Value;
            txtFinalContrato.Text = fechaFinalContrato.ToString();
        }

        private void dateTimePicker11_ValueChanged(object sender, EventArgs e)
        {
            DateTime fechaRealTrabajo = dateTimePicker11.Value;
            txtRealTrabajo.Text = fechaRealTrabajo.ToString();
        }

        private void dateTimePicker12_ValueChanged(object sender, EventArgs e)
        {
            DateTime fechaFinalPrev = dateTimePicker12.Value;
            txtFinalPrevista.Text = fechaFinalPrev.ToString();
        }

        private void FormFact_Load(object sender, EventArgs e)
        {
            CargaInicio();
            //btnAgregar.Enabled = false;
        }

        private void txtRuc1_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroMes = new List<TabFactura>();
            listaFacturaFiltroMes = listaFactura.Where(item => item.ruc_cl.Contains(txtRuc1.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltroMes;

        }

        private void txtTipoFactura_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltroTipoFactura = new List<TabFactura>();
            listaFacturaFiltroTipoFactura = listaFactura.Where(item => item.tipofactura.Contains(txtTipoFactura.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listaFacturaFiltroTipoFactura;
        }

        private void txtNumbFact_TextChanged(object sender, EventArgs e)
        {
            List<TabFactura> listaFacturaFiltrNumerFactura = new List<TabFactura>();
            listaFacturaFiltrNumerFactura = listaFactura.Where(item => item.numFactura.Contains(txtNumbFact.Text)).ToList();
            dataGridView1.DataSource = listaFacturaFiltrNumerFactura;
        }
    }

}