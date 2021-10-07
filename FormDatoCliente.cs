using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static facturacion.Entidades;
using static facturacion.FormSistema;

namespace facturacion
{
    public partial class FormDatoCliente : Form
    {

        //CADENA DE INSTANCIACION DE LA BASE DE DATOS
        String rutaAcces = Entidades.Variablesglobales.rutaAccess;
        String microsolftAccessDatabaseConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = ";


        String selectDataFromMSAccessDatabaseQuery = "SELECT * FROM TablaCliente ";
        String InsertDataIntoMicrosoftAccessDatabase = "INSERT INTO TablaCliente (ruc_cl,nomb_emp,direccion,pagWeb,telefonoFijo,correoEmpresa,contacto,cargo,celular,correo) VALUES (?,?,?,?,?,?,?,?,?,?)";
        String UpdateMicrosoftAccessDataQuery = "UPDATE TablaCliente SET ruc_cl = ? ,nomb_emp = ? ,direccion = ?,pagWeb= ?,telefonoFijo = ?,correoEmpresa= ?,contacto =?,cargo=?,celular=?,correo=? WHERE Id = ?";
        String DeleteMicrosoftAccessDataQuery = "DELETE * FROM TablaCliente WHERE Id = ?";

        OleDbConnection AccessDatabaseeConecctionStringOleDbConnection = null;

        List<TabClientes> DatosCliente = new List<TabClientes>();

        public FormDatoCliente()
        {
            InitializeComponent();
        }


        public void OpenConexion()
        {
            AccessDatabaseeConecctionStringOleDbConnection = new OleDbConnection(microsolftAccessDatabaseConnectionString + rutaAcces);
            if (AccessDatabaseeConecctionStringOleDbConnection.State == ConnectionState.Closed)
            {
                AccessDatabaseeConecctionStringOleDbConnection.Open();
            }
        }

        public void CloseConexion()
        {
            AccessDatabaseeConecctionStringOleDbConnection = new OleDbConnection(microsolftAccessDatabaseConnectionString + rutaAcces);
            if (AccessDatabaseeConecctionStringOleDbConnection.State == ConnectionState.Open)
            {
                AccessDatabaseeConecctionStringOleDbConnection.Close();
            }
        }

        public void InciarDatos()
        {
            OpenConexion();
            DatosCliente = PopulateDataGridViewFromMicrosoftAccessDatabase();
            dataGridView1.DataSource = DatosCliente;
            //DarFormatoCabecera();
        }

        public List<TabClientes> PopulateDataGridViewFromMicrosoftAccessDatabase() //FUNCION PARA LEER LOS DATOS DE LA BASE DE DATOS
        {
            //control k + control s  : sacar el TRY CATH
            //creamos la lista en una instania de lista de la tabla Factura
            List<TabClientes> lista = new List<TabClientes>();
            try
            {
                //// borrar filas de cuadrícula de datos antes de cargar datos de accesos de microsoft
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.DataSource = null;
                }
                OleDbCommand populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(selectDataFromMSAccessDatabaseQuery, AccessDatabaseeConecctionStringOleDbConnection);
                //abrimos la conexion a base de datos
                OpenConexion();

                //Es el nombre de la variable , se crea para poder instanciar y poder llenar los datos y
                //para que se guarden de forma ordenada se crea un lista
                TabClientes objeto;

                OleDbDataReader reader = populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                // bucle para leer los datos de Microsoft Access Database
                while (reader.Read())
                {
                    objeto = new TabClientes();
                    objeto.Id = (int)reader[0];
                    objeto.ruc_cl = reader[1].ToString();
                    objeto.nomb_emp = reader[2].ToString();
                    objeto.direccion = reader[3].ToString();
                    objeto.pagWeb = reader[4].ToString();
                    objeto.telefonoFijo = reader[5].ToString();
                    objeto.correoEmpresa = reader[6].ToString();
                    objeto.contacto = reader[7].ToString();
                    objeto.cargo = reader[8].ToString();
                    objeto.celular = reader[9].ToString();
                    objeto.correo = reader[10].ToString();
                    //valores en fechas
                    /*if (Information.IsDBNull(reader[10]) != true)
                    {
                        objeto.fechaCobroDetra = (DateTime)reader[10];
                    }*/
                    lista.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return lista;
            }
            finally
            {
                //cerramos la conexion con la base de datos
                CloseConexion();
            }
            return lista;
        }
       

        public void DarFormatoCabecera()
        {
            //DataGridViewColumn columna0 = dataGridView1.Columns[0];
            //columna0.HeaderText = "ID";
            //columna0.Width = 30;

            //DataGridViewColumn columna = dataGridView1.Columns[2];
            //columna.HeaderText = "Nombre de la Empresa";
            //columna.Width = 200;

        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            txtIdCli.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtRuc_Cliente.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtNombEmp.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            txtDireccion.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtPagWeb.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            txtTelefonoFijo.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            txtCorreoEmpresa.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            txtContacto.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            txtCargo.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            txtCelular.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            txtCorreo.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
        }

        private void limpiar()
        {

            txtRuc_Cliente.Text = String.Empty;
            txtNombEmp.Text = String.Empty;
            txtDireccion.Text = String.Empty;
            txtPagWeb.Text = String.Empty;
            txtTelefonoFijo.Text = String.Empty;
            txtCorreoEmpresa.Text = String.Empty;
            txtContacto.Text = String.Empty;
            txtCargo.Text = String.Empty;
            txtCelular.Text = String.Empty;
            txtCorreo.Text = String.Empty;
        }
        
      

        private void btnDeconectDDBB_Click(object sender, EventArgs e)
        {
            CloseConexion();
            limpiar();
            dataGridView1.DataSource = null;
            lblRuta.Text = "";
            MessageBox.Show("Desconectado de la Base de datos .......");
        }

        private void txtFilruc_TextChanged(object sender, EventArgs e)
        {
            List<TabClientes> listaClientesFiltroCliente = new List<TabClientes>();
            listaClientesFiltroCliente = DatosCliente.Where(item => item.ruc_cl.Contains(txtFilruc.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listaClientesFiltroCliente;
        }

        private void txtFilnomEmpr_TextChanged(object sender, EventArgs e)
        {
            List<TabClientes> listarNombreEmpresaFil = new List<TabClientes>();
            listarNombreEmpresaFil = DatosCliente.Where(item => item.nomb_emp.Contains(txtFilnomEmpr.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listarNombreEmpresaFil;
        }

        private void txtContact_TextChanged(object sender, EventArgs e)
        {
            List<TabClientes> listarContacFiltrar = new List<TabClientes>();
            listarContacFiltrar = DatosCliente.Where(item => item.contacto.Contains(txtFilContacto.Text.ToUpper())).ToList();
            dataGridView1.DataSource = listarContacFiltrar;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            OpenConexion();
            OleDbCommand insertDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(InsertDataIntoMicrosoftAccessDatabase, AccessDatabaseeConecctionStringOleDbConnection);
            if (txtNombEmp.Text == String.Empty || txtDireccion.Text == String.Empty || txtContacto.Text == String.Empty || txtTelefonoFijo.Text == String.Empty || txtCorreoEmpresa.Text == String.Empty || txtPagWeb.Text == String.Empty || txtCargo.Text == string.Empty)

            {
                MessageBox.Show("Verificar que uno o más campos vacíos estén llenos............");

            }
            try
            {
                //los nombres de los campos debe ser correspondientes con la base de datos Access

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = txtRuc_Cliente.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmp.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("direccion", OleDbType.VarChar).Value = txtDireccion.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("pagWeb", OleDbType.VarChar).Value = txtPagWeb.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("telefonoFijo", OleDbType.VarChar).Value = txtTelefonoFijo.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("correoEmpresa", OleDbType.VarChar).Value = txtCorreoEmpresa.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("contacto", OleDbType.VarChar).Value = txtContacto.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cargo", OleDbType.VarChar).Value = txtCargo.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("celular", OleDbType.VarChar).Value = txtCelular.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("correo", OleDbType.VarChar).Value = txtCorreoEmpresa.Text;

                int DataInsert = insertDataIntoMSAccessDataBaseOleDbCommand.ExecuteNonQuery();
                if (DataInsert > 0)
                {
                    //MessageBox.Show("Registro Exitosoo ¡¡¡");
                    MessageBox.Show("Registro Exitoso", "Registro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Refrescar el Datagrid al insertar un registro
                    DatosCliente = PopulateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = DatosCliente;
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
                CloseConexion();
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {

            if (String.IsNullOrEmpty(txtIdCli.Text))
            {
                MessageBox.Show("Hacer click en uno de los registros , que desee actualizar");
            }
            else
            {
                //abrir conexion
                OpenConexion();
                OleDbCommand UpdateDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(UpdateMicrosoftAccessDataQuery, AccessDatabaseeConecctionStringOleDbConnection);
                //condicional para que evalue si todos estan llenos

                try
                {
                    //los nombres de los campos debe ser correspondientes con la base de datos Access
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("ruc_cl", OleDbType.VarChar).Value = txtRuc_Cliente.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomb_emp", OleDbType.VarChar).Value = txtNombEmp.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("direccion", OleDbType.VarChar).Value = txtDireccion.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("pagWeb", OleDbType.VarChar).Value = txtPagWeb.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("telefonoFijo", OleDbType.VarChar).Value = txtTelefonoFijo.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("correoEmpresa", OleDbType.VarChar).Value = txtCorreoEmpresa.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("contacto", OleDbType.VarChar).Value = txtContacto.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cargo", OleDbType.VarChar).Value = txtCargo.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("celular", OleDbType.VarChar).Value = txtCelular.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("correo", OleDbType.VarChar).Value = txtCorreoEmpresa.Text;
                    UpdateDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Id", OleDbType.Integer).Value = Convert.ToInt16(txtIdCli.Text);

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
                    DatosCliente = PopulateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = DatosCliente;
                    //DarFormatoCabecera();
                    CloseConexion();
                }


            }
        }

        private void btnErase_Click(object sender, EventArgs e)
        {
            OpenConexion();

            try
            {
                if (String.IsNullOrEmpty(txtIdCli.Text))
                {
                    MessageBox.Show("Campo ID Vacio , seleccionar un fila de registro");
                }
                else
                {
                    OleDbCommand DeleteMicrosoftAccessDataQueryOleDbCommand = new OleDbCommand(DeleteMicrosoftAccessDataQuery, AccessDatabaseeConecctionStringOleDbConnection);
                    DeleteMicrosoftAccessDataQueryOleDbCommand.Parameters.AddWithValue("Id", OleDbType.Integer).Value = Convert.ToUInt16(txtIdCli.Text);
                    //abrir la coneexion


                    int DeleteMicrosoftAccessData = DeleteMicrosoftAccessDataQueryOleDbCommand.ExecuteNonQuery();
                    if (DeleteMicrosoftAccessData > 0)
                    {
                        MessageBox.Show("Registro Borrado");
                        //Se debe refrescar el datagrid antes de actualizar los datos eliminados de la base de datos de ACCESSS
                        //limpiar los capos
                        limpiar();
                        DatosCliente = PopulateDataGridViewFromMicrosoftAccessDatabase();
                        dataGridView1.DataSource = DatosCliente;
                        //DarFormatoCabecera();
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
                CloseConexion();
            }
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            limpiar();
        }

        private void btnExpor_Click(object sender, EventArgs e)
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
                    InciarDatos();
                    //MessageBox.Show("Archivo correcto");
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Primero Necesita Realizar la Conexion a la Base de datos", "Verificar Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();

            }

        }

        private void FormDatoCliente_Load(object sender, EventArgs e)
        {
            CargaInicio();
        }
    }
}
