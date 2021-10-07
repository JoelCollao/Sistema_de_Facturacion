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

namespace facturacion
{
    public partial class FormOferta : Form
    {
        
        /// //////////////////////////////////////////----- CONEXION BASE DE DATOS---------------------//////////////////////////////////7
       
        String rutaAcces = Entidades.Variablesglobales.rutaAccess;
        String microsolftAccessDatabaseConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = ";  //LEER DE CUALQUIER DISCO C: D: F: etc

        ////////////////////////////////////////////////------------- CRUD A LA BASE DE DATOS -----------------------////////////////////////////
        
        String selectDataFromMSAccessDatabaseQuery = "SELECT * FROM TablaOferta";
        String InsertDataIntoMicrosoftAccessDatabase = "INSERT INTO TablaOferta (codOferta, fecha, anho, mes, dia, nomContac, telefono, emailClient, nomEmpre, proyecto ,titConcep1, titConcep2, titConcep3,cant1, cant2, cant3, cant2a, cant3a, cant3b, cant3c, cant3d , cant3e,und1 ,und2 ,und2a ,und3 ,und3a ,und3b ,und3c ,und3d ,und3e ,descripConcep1 ,descripConcep2 ,descripConcep2a ,descripConcep3 ,descripConcep3a ,descripConcep3b ,descripConcep3c ,descripConcep3d ,descripConcep3e,import1 ,import2 ,import2a ,import3 ,import3a ,import3b ,import3c ,import3d ,import3e ,condOfert1 ,condOfert2 ,condOfert3 , condicion0 ,condicion1 ,condicion2 ,condicion3 ,Total1, Total2 ,Total3 ,Total4 ,Total5 ,Total6,Total7 ,Total8 ,Total9 ,Subtotal ,igv18 ,SumaTotal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        String UpdateMicrosoftAccessDataQuery = "UPDATE TablaOferta SET codOferta= ?, fecha= ?, anho= ?, mes= ?, dia= ?, nomContac= ?, telefono= ?, emailClient= ?, nomEmpre= ?, proyecto = ?,titConcep1= ?, titConcep2= ?, titConcep3= ?,cant1= ?, cant2= ?, cant3= ?, cant2a= ?, cant3a= ?, cant3b= ?, cant3c= ?, cant3d = ?, cant3e= ?,und1 = ?,und2 = ?,und2a = ?,und3 = ?,und3a = ?,und3b = ?,und3c = ?,und3d = ?,und3e = ?,descripConcep1 = ?,descripConcep2 = ?,descripConcep2a = ?,descripConcep3 = ?,descripConcep3a = ?,descripConcep3b = ?,descripConcep3c = ?,descripConcep3d = ?,descripConcep3e= ?,import1 = ?,import2 = ?,import2a = ?,import3 = ?,import3a = ?,import3b = ?,import3c = ?,import3d = ?,import3e = ?,condOfert1 = ?,condOfert2 = ?,condOfert3 = ?, condicion0 = ?,condicion1 = ?,condicion2 = ?,condicion3 = ?,Total1= ?, Total2 = ?,Total3 = ?,Total4 = ?,Total5 = ?,Total6= ?,Total7 = ?,Total8 = ?,Total9 = ?,Subtotal = ?,igv18 = ?,SumaTotal=? WHERE idoferta = ?";
        String DeleteMicrosoftAccessDataQuery = "DELETE * FROM TablaOferta WHERE idoferta = ?";



        /// /////////////////////////////////---- CONSULTA A LA TABLA DE TITULO DE CONCEPTO DE LA BASE DE DATOS ---------------/////////////////////7

        string QueryTablaTituConcepto = "SELECT tituConcept FROM TablaTituloConceptoOferta";


        //////////////// -------------------- CONSULTA A LA TABLA DE DESCRIPCION DEL CONCEPTO  -----------------------------------------//////////////////

        string QueryTablaDescripcionConcepto = "SELECT descrConcept FROM TablaDescripConcepto";



        ////////////// ----------------------------- CONSULTA A LA TABLA DE CONDICIONES , DE LA BASE DE DATOS ------------------////////////////////////

        string QueryTablaCondiciones = "SELECT condici FROM TablaCondiciones";



        ////////////  -------------------------- CONSULTA A LA TABLA DE CONDICONES ----------------------------------- //////////////////////////////////

        string QueryTablaCondicionesOferta = "SELECT condOferta FROM TablaCondicionesOferta";


        OleDbConnection AccessDatabaseConecctionStringOleConecction = null;

        List<TablaOfertas> ListaDatosOferta = new List<TablaOfertas>();

        ////////// --------------- LISTA PARA TRAER EN EL COMBOBOX LOS TITULOS DE LOS CONCEPTOS -------------------//////////////////////////////////////////////////
        
        List<TablaTituloConceptoOferta> ListaTituloConcepto = new List<TablaTituloConceptoOferta>();
        List<TablaTituloConceptoOferta2> ListaTituloConcepto2 = new List<TablaTituloConceptoOferta2>();
        List<TablaTituloConceptoOferta3> ListaTituloConcepto3 = new List<TablaTituloConceptoOferta3>();


        /// //////////////-----------------LISTA PARA TRAER AL COMBOBOX LA DESCRIPCION DE LOS CONCEPTOS ----------------/////////////////////////////////
        List<TablaDescripcionTitulo1> ListaDescripcionTitulo1 = new List<TablaDescripcionTitulo1>();
        List<TablaDescripcionTitulo2> ListaDescripcionTitulo2 = new List<TablaDescripcionTitulo2>();
        List<TablaDescripcionTitulo3> ListaDescripcionTitulo3 = new List<TablaDescripcionTitulo3>();
        List<TablaDescripcionTitulo4> ListaDescripcionTitulo4 = new List<TablaDescripcionTitulo4>();
        List<TablaDescripcionTitulo5> ListaDescripcionTitulo5 = new List<TablaDescripcionTitulo5>();
        List<TablaDescripcionTitulo6> ListaDescripcionTitulo6 = new List<TablaDescripcionTitulo6>();
        List<TablaDescripcionTitulo7> ListaDescripcionTitulo7 = new List<TablaDescripcionTitulo7>();
        List<TablaDescripcionTitulo8> ListaDescripcionTitulo8 = new List<TablaDescripcionTitulo8>();
        List<TablaDescripcionTitulo9> ListaDescripcionTitulo9 = new List<TablaDescripcionTitulo9>();


        //////////////------------ LISTA PARA TRAER AL COMBOBOX LAS CONDICIONES DE LA OFERTA --------------------------////////////////////////////////////

        List<TabladeCondiciones1> ListaCondiciones1 = new List<TabladeCondiciones1>();
        List<TabladeCondiciones2> ListaCondiciones2 = new List<TabladeCondiciones2>();
        List<TabladeCondiciones3> ListaCondiciones3 = new List<TabladeCondiciones3>();
        List<TabladeCondiciones4> ListaCondiciones4 = new List<TabladeCondiciones4>();


        ///////////// ----------- LISTA PARA TRAER AL COMBOBOX LAS CONDICIONES - OFERTAS ---------------------- //////////////////////////

        List<TabladeCondicionOferta1> ListaCondicionesOferta1 = new List<TabladeCondicionOferta1>();
        List<TabladeCondicionOferta2> ListaCondicionesOferta2 = new List<TabladeCondicionOferta2>();
        List<TabladeCondicionOferta3> ListaCondicionesOferta3 = new List<TabladeCondicionOferta3>();

        public FormOferta()
        {
            InitializeComponent();
        }
        public void Openconectar()
        {
           
            AccessDatabaseConecctionStringOleConecction = new OleDbConnection(microsolftAccessDatabaseConnectionString + rutaAcces);
            if (AccessDatabaseConecctionStringOleConecction.State == ConnectionState.Closed)
            {
                AccessDatabaseConecctionStringOleConecction.Open();
            }

        }

        ////////////////////////////////////////////////////////////   ------    ////////////////////////////////////////////////////////////////////////
        public void CerrarConexion()
        {
            AccessDatabaseConecctionStringOleConecction = new OleDbConnection(microsolftAccessDatabaseConnectionString + rutaAcces);
            if (AccessDatabaseConecctionStringOleConecction.State == ConnectionState.Open)
            {
                AccessDatabaseConecctionStringOleConecction.Close();
            }
        }

        ////////////////////////////////////////////////////  ----- AGREGAR LOS DATOS DE LA TABLA DE OFERTA ----////////////////////////////////////////


        public List<TablaOfertas> PopulateDataGridViewFromMicrosoftAccessDatabase() 
        {
           
            List<TablaOfertas> lista = new List<TablaOfertas>();
            try
            {
             
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.DataSource = null;
                }
                OleDbCommand populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(selectDataFromMSAccessDatabaseQuery, AccessDatabaseConecctionStringOleConecction);
               
                Openconectar();
                TablaOfertas objeto;
                OleDbDataReader reader = populateDataGridViewFromMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                
                while (reader.Read())
                {
                    objeto = new TablaOfertas();
                    objeto.idoferta = (int)reader[0];
                    objeto.codOferta = reader[1].ToString();

                   if (Information.IsDBNull(reader[2]) != true)
                   {
                      objeto.fecha = (DateTime)reader[2];
                   }
                    objeto.anho = reader[3].ToString();
                    objeto.mes = reader[4].ToString();
                    objeto.dia = reader[5].ToString();
                    objeto.nomContac = reader[6].ToString();
                    objeto.telefono = reader[7].ToString();
                    objeto.emailClient = reader[8].ToString();
                    objeto.nomEmpre = reader[9].ToString();
                    objeto.proyecto = reader[10].ToString();
                    objeto.titConcep1 = reader[11].ToString();
                    objeto.titConcep2 = reader[12].ToString();
                    objeto.titConcep3 = reader[13].ToString();

                    ///////-----------------------------    CANTIDAD     ------------------------------------------------------------/////////////////////

                    if (Information.IsDBNull(reader[14]) != true)
                    {
                        objeto.cant1 = (int)reader[14];
                    }

                    if (Information.IsDBNull(reader[15]) != true)
                    {
                        objeto.cant2 = (int)reader[15];
                    }

                    if (Information.IsDBNull(reader[16]) != true)
                    {
                        objeto.cant3 = (int)reader[16];
                    }

                    if (Information.IsDBNull(reader[17]) != true)
                    {
                        objeto.cant2a = (int)reader[17];
                    }

                    if (Information.IsDBNull(reader[18]) != true)
                    {
                        objeto.cant3a = (int)reader[18];
                    }

                    if (Information.IsDBNull(reader[19]) != true)
                    {
                        objeto.cant3b = (int)reader[19];
                    }

                    if (Information.IsDBNull(reader[20]) != true)
                    {
                        objeto.cant3c = (int)reader[20];
                    }

                    if (Information.IsDBNull(reader[21]) != true)
                    {
                        objeto.cant3d = (int)reader[21];
                    }

                    if (Information.IsDBNull(reader[22]) != true)
                    {
                        objeto.cant3e = (int)reader[22];
                    }
              
                    ///////-------- UNIDADES ------------------------------------------------------------------/////////////////////

                    objeto.und1 = reader[23].ToString();
                    objeto.und2 = reader[24].ToString();
                    objeto.und2a = reader[25].ToString();
                    objeto.und3 = reader[26].ToString();
                    objeto.und3a = reader[27].ToString();
                    objeto.und3b = reader[28].ToString();
                    objeto.und3c = reader[29].ToString();
                    objeto.und3d = reader[30].ToString();
                    objeto.und3e = reader[31].ToString();

                    
                    ///////-------- DESCRIPCION -------------------------------------------------------------------////////////////////////

                    objeto.descripConcep1 = reader[32].ToString();
                    objeto.descripConcep2 = reader[33].ToString();
                    objeto.descripConcep2a = reader[34].ToString();
                    objeto.descripConcep3 = reader[35].ToString();
                    objeto.descripConcep3a = reader[36].ToString();
                    objeto.descripConcep3b = reader[37].ToString();
                    objeto.descripConcep3c = reader[38].ToString();
                    objeto.descripConcep3d = reader[39].ToString();
                    objeto.descripConcep3e = reader[40].ToString();

                    ///////-------- IMPORTE -----------------------------------------------------------//////////////////////////

                    if (Information.IsDBNull(reader[41])!= true)
                    {
                        objeto.import1 = float.Parse((reader[41]).ToString());
                    }
                    
                  
                    if (Information.IsDBNull(reader[42])!= true)
                    {
                        objeto.import2 = float.Parse((reader[42]).ToString());
                    }

                    if (Information.IsDBNull(reader[43])!=true)
                    {
                        objeto.import2a = float.Parse((reader[43]).ToString());
                    }

                    if (Information.IsDBNull(reader[44]) != true)
                    {
                        objeto.import3 = float.Parse((reader[44]).ToString());
                    }

                    if (Information.IsDBNull(reader[45])!=true)
                    {
                        objeto.import3a = float.Parse((reader[45]).ToString());
                    }

                    if(Information.IsDBNull(reader[46]) != true)
                    {
                        objeto.import3b = float.Parse((reader[46]).ToString());
                    }

                    if (Information.IsDBNull(reader[47])!=true)
                    {
                        objeto.import3c = float.Parse((reader[47]).ToString());
                    }

                    if (Information.IsDBNull(reader[48]) != true)
                    {
                        objeto.import3d = float.Parse((reader[48]).ToString());
                    }
                    
                    if (Information.IsDBNull(reader[49])!= true)
                    {
                        objeto.import3e = float.Parse((reader[49]).ToString());
                    }
                    

                    ///////-------- CONDICION - OFERTA ----------------------------------------------------//////////////////////////

                    objeto.condOfert1 = reader[50].ToString();
                    objeto.condOfert2 = reader[51].ToString();
                    objeto.condOfert3 = reader[52].ToString();


                    ///////-------- CONDICION ----------------------------------------------------------------------//////////////////

                    objeto.condicion0 = reader[53].ToString();
                    objeto.condicion1 = reader[54].ToString();
                    objeto.condicion2 = reader[55].ToString();
                    objeto.condicion3 = reader[56].ToString();


                    ///////----------- TOTAL --------------------------------------------------------------------------/////////////////


                    if (Information.IsDBNull(reader[57]) != true)
                    {
                        objeto.Total1 = float.Parse(reader[57].ToString());
                    }

                    if (Information.IsDBNull(reader[58]) != true)
                    {
                        objeto.Total2 = float.Parse(reader[58].ToString());
                    }

                    if (Information.IsDBNull(reader[59]) != true)
                    {
                        objeto.Total3 = float.Parse(reader[59].ToString());
                    }

                    if (Information.IsDBNull(reader[60]) != true)
                    {
                        objeto.Total4 = float.Parse(reader[60].ToString());
                    }

                    if (Information.IsDBNull(reader[61]) != true)
                    {
                        objeto.Total5 = float.Parse(reader[61].ToString());
                    }

                    if (Information.IsDBNull(reader[62]) != true)
                    {
                        objeto.Total6 = float.Parse(reader[62].ToString());
                    }

                    if (Information.IsDBNull(reader[63]) != true)
                    {
                        objeto.Total7 = float.Parse(reader[63].ToString());
                    }

                    if (Information.IsDBNull(reader[64]) != true)
                    {
                        objeto.Total8 = float.Parse(reader[64].ToString());
                    }

                    if (Information.IsDBNull(reader[65]) != true)
                    {
                        objeto.Total9 = float.Parse(reader[65].ToString());
                    }

                    if (Information.IsDBNull(reader[66]) != true)
                    {
                        objeto.Subtotal = float.Parse(reader[66].ToString());
                    }

                    if (Information.IsDBNull(reader[67]) != true)
                    {
                        objeto.igv18 = float.Parse(reader[67].ToString());
                    }

                    if (Information.IsDBNull(reader[68]) != true)
                    {
                        objeto.SumaTotal = float.Parse(reader[68].ToString());
                    }

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
                CerrarConexion();
            }
            return lista;
        }

        //////////////////////////////////////////////    ---------------    DATETIMEPICKER   --------------    ////////////////////////////////////////////////////////////

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime fecha1 = dateTimePicker1.Value;
            txtFecha.Text = fecha1.ToString();
            txtanho.Text = fecha1.ToString("yyyy");
            txtMes.Text = fecha1.ToString("MM");
            txtDia.Text = fecha1.ToString("dd");
        }

        ////////////////////////// ------ LISTA LOS DATOS DE LA TABLA DE TITULO DE CONCEPTO --- AQUI ESTA LAS 3 LLENADOS DE CONCEPTO ----- ////////////////////////////////////////////

        public List<TablaTituloConceptoOferta> PopulatedAddComboboxConceptoTitulo()
        {
            List<TablaTituloConceptoOferta> ListaTituloConcepto = new List<TablaTituloConceptoOferta>();
            try
            {
                if (cbTituConcep1.Items.Count > 0)
                {
                    cbTituConcep1.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaTituConcepto, AccessDatabaseConecctionStringOleConecction); 
                Openconectar();
                TablaTituloConceptoOferta objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaTituloConceptoOferta();
                    objeto.tituConcept = reader[0].ToString();
                    ListaTituloConcepto.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaTituloConcepto;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaTituloConcepto;
        }

        ///----------------------------------------------------------------------------- LISTA 2 DE TITULO DE CONCEPTO PARA COMBOXBOX -----------------------

        public List<TablaTituloConceptoOferta2> PopulatedAddComboboxConceptoTitulo2()
        {
            List<TablaTituloConceptoOferta2> ListaTituloConcepto2 = new List<TablaTituloConceptoOferta2>();
            try
            {
                if (cbTituConcep2.Items.Count > 0)
                {
                    cbTituConcep2.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaTituConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaTituloConceptoOferta2 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaTituloConceptoOferta2();
                    objeto.tituConcept = reader[0].ToString();
                    ListaTituloConcepto2.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaTituloConcepto2;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaTituloConcepto2;
        }

      
        ///----------------------------------------------------------------------------- LISTA 3 DE TITULO DE CONCEPTO PARA COMBOXBOX -----------------------

        public List<TablaTituloConceptoOferta3> PopulatedAddComboboxConceptoTitulo3()
        {
            List<TablaTituloConceptoOferta3> ListaTituloConcepto3 = new List<TablaTituloConceptoOferta3>();
            try
            {
                if (cbTituConcep03.Items.Count > 0)
                {
                    cbTituConcep03.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaTituConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaTituloConceptoOferta3 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaTituloConceptoOferta3();
                    objeto.tituConcept = reader[0].ToString();
                    ListaTituloConcepto3.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaTituloConcepto3;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaTituloConcepto3;
        }

        ///////////////////////////--------------------------- FUNCIONES PARA LEER TABLA DE DESCRIPCION DE TITULO , EN EL COMBOBOX ------//////

        public List<TablaDescripcionTitulo1> PopulatedAddComboboxDescripcionTitulo1()
        {
            List<TablaDescripcionTitulo1> ListaDescripcionTitulo1 = new List<TablaDescripcionTitulo1>();
            try
            {
                if (cbDescripConcep1.Items.Count > 0)
                {
                    cbDescripConcep1.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo1 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo1();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo1.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo1;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo1;
        }

        /////////////////////////////////---------------------------- DESCRIPCION DE TITUTLO 2 -----------------------------------------------------------

        public List<TablaDescripcionTitulo2> PopulatedAddComboboxDescripcionTitulo2()
        {
            List<TablaDescripcionTitulo2> ListaDescripcionTitulo2 = new List<TablaDescripcionTitulo2>();
            try
            {
                if (cbDescripConcep2.Items.Count > 0)
                {
                    cbDescripConcep2.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo2 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo2();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo2.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo2;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo2;
        }

       
        //////////////////////////////////////////// ---------------------- DESCRIPCION DE TITULO 3 -----------------------/////////////////////////////////////////
       
        public List<TablaDescripcionTitulo3> PopulatedAddComboboxDescripcionTitulo3()
        {
            List<TablaDescripcionTitulo3> ListaDescripcionTitulo3 = new List<TablaDescripcionTitulo3>();
            try
            {
                if (cbDescripConcep3.Items.Count > 0)
                {
                    cbDescripConcep3.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo3 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo3();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo3.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo3;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo3;
        }


        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 4 -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo4> PopulatedAddComboboxDescripcionTitulo4()
        {
            List<TablaDescripcionTitulo4> ListaDescripcionTitulo4 = new List<TablaDescripcionTitulo4>();
            try
            {
                if (cbDescripConcep4.Items.Count > 0)
                {
                    cbDescripConcep4.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo4 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo4();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo4.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo4;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo4;
        }


        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 5 -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo5> PopulatedAddComboboxDescripcionTitulo5()
        {
            List<TablaDescripcionTitulo5> ListaDescripcionTitulo5 = new List<TablaDescripcionTitulo5>();
            try
            {
                if (cbDescripConcep5.Items.Count > 0)
                {
                    cbDescripConcep5.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo5 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo5();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo5.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo5;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo5;
        }


        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 6  -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo6> PopulatedAddComboboxDescripcionTitulo6()
        {
            List<TablaDescripcionTitulo6> ListaDescripcionTitulo6 = new List<TablaDescripcionTitulo6>();
            try
            {
                if (cbDescripConcep6.Items.Count > 0)
                {
                    cbDescripConcep6.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo6 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo6();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo6.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo6;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo6;
        }

        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 7  -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo7> PopulatedAddComboboxDescripcionTitulo7()
        {
            List<TablaDescripcionTitulo7> ListaDescripcionTitulo7 = new List<TablaDescripcionTitulo7>();
            try
            {
                if (cbDescripConcep7.Items.Count > 0)
                {
                    cbDescripConcep7.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo7 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo7();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo7.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo7;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo7;
        }


        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 8  -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo8> PopulatedAddComboboxDescripcionTitulo8()
        {
            List<TablaDescripcionTitulo8> ListaDescripcionTitulo8 = new List<TablaDescripcionTitulo8>();
            try
            {
                if (cbDescripConcep7.Items.Count > 0)
                {
                    cbDescripConcep7.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo8 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo8();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo8.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo8;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo8;
        }


        ///////////////////////////////////////////-------------------------- DESCRIPCION DE TITULO 9  -----------------------//////////////////////////////////////

        public List<TablaDescripcionTitulo9> PopulatedAddComboboxDescripcionTitulo9()
        {
            List<TablaDescripcionTitulo9> ListaDescripcionTitulo9 = new List<TablaDescripcionTitulo9>();
            try
            {
                if (cbDescripConcep9.Items.Count > 0)
                {
                    cbDescripConcep9.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaDescripcionConcepto, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TablaDescripcionTitulo9 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TablaDescripcionTitulo9();
                    objeto.descrConcept = reader[0].ToString();
                    ListaDescripcionTitulo9.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaDescripcionTitulo9;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaDescripcionTitulo9;
        }

        /// ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// ////////////////////////////////-------------------   LLENADO DE TABLA DE CONDICIONES  1  ------ //////////////////////////////////////////
        /// //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


        public List<TabladeCondiciones1> PopulatedAddComboboxCondiciones1()
        {
            List<TabladeCondiciones1> ListaCondiciones1 = new List<TabladeCondiciones1>();
            try
            {
                if (cbCondiciones1.Items.Count > 0)
                {
                    cbCondiciones1.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondiciones, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondiciones1 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondiciones1();
                    objeto.condici = reader[0].ToString();
                    ListaCondiciones1.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondiciones1;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondiciones1;
        }


        ////////////////////////////////////-------------------   LLENADO DE TABLA DE CONDICIONES  2  ------ //////////////////////////////////////////

        public List<TabladeCondiciones2> PopulatedAddComboboxCondiciones2()
        {
            List<TabladeCondiciones2> ListaCondiciones2 = new List<TabladeCondiciones2>();
            try
            {
                if (cbCondiciones2.Items.Count > 0)
                {
                    cbCondiciones2.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondiciones, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondiciones2 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondiciones2();
                    objeto.condici = reader[0].ToString();
                    ListaCondiciones2.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondiciones2;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondiciones2;
        }

        ////////////////////////////////////-------------------   LLENADO DE TABLA DE CONDICIONES  3  ------ //////////////////////////////////////////

        public List<TabladeCondiciones3> PopulatedAddComboboxCondiciones3()
        {
            List<TabladeCondiciones3> ListaCondiciones3 = new List<TabladeCondiciones3>();
            try
            {
                if (cbCondiciones3.Items.Count > 0)
                {
                    cbCondiciones3.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondiciones, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondiciones3 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondiciones3();
                    objeto.condici = reader[0].ToString();
                    ListaCondiciones3.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondiciones3;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondiciones3;
        }


        ////////////////////////////////////-------------------   LLENADO DE TABLA DE CONDICIONES  4  ------ //////////////////////////////////////////

        public List<TabladeCondiciones4> PopulatedAddComboboxCondiciones4()
        {
            List<TabladeCondiciones4> ListaCondiciones4 = new List<TabladeCondiciones4>();
            try
            {
                if (cbCondiciones4.Items.Count > 0)
                {
                    cbCondiciones4.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondiciones, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondiciones4 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondiciones4();
                    objeto.condici = reader[0].ToString();
                    ListaCondiciones4.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondiciones4;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondiciones4;
        }


        /////////////////////////////////////////  --------- TABLA DE CONDICIONES - OFERTAS 1 --------------///////////////////////////////////////

        public List<TabladeCondicionOferta1> PopulatedAddComboboxCondicion_Oferta1()
        {
            List<TabladeCondicionOferta1> ListaCondicionesOferta = new List<TabladeCondicionOferta1>();
            try
            {
                if (cbCondiOferta1.Items.Count > 0)
                {
                    cbCondiOferta1.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondicionesOferta, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondicionOferta1 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondicionOferta1();
                    objeto.condOferta = reader[0].ToString();
                    ListaCondicionesOferta.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondicionesOferta;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondicionesOferta;
        }


        /////////////////////////////////////////  --------- TABLA DE CONDICIONES - OFERTAS 2 --------------///////////////////////////////////////

        public List<TabladeCondicionOferta2> PopulatedAddComboboxCondicion_Oferta2()
        {
            List<TabladeCondicionOferta2> ListaCondicionesOferta2 = new List<TabladeCondicionOferta2>();
            try
            {
                if (cbCondiOferta2.Items.Count > 0)
                {
                    cbCondiOferta2.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondicionesOferta, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondicionOferta2 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondicionOferta2();
                    objeto.condOferta = reader[0].ToString();
                    ListaCondicionesOferta2.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondicionesOferta2;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondicionesOferta2;
        }

        /////////////////////////////////////////  --------- TABLA DE CONDICIONES - OFERTAS 3 --------------///////////////////////////////////////

        public List<TabladeCondicionOferta3> PopulatedAddComboboxCondicion_Oferta3()
        {
            List<TabladeCondicionOferta3> ListaCondicionesOferta3 = new List<TabladeCondicionOferta3>();
            try
            {
                if (cbCondiOferta3.Items.Count > 0)
                {
                    cbCondiOferta3.DataSource = null;
                }
                OleDbCommand TableFacturaMicrosoftAccessDatabaseOleDbCommand = new OleDbCommand(QueryTablaCondicionesOferta, AccessDatabaseConecctionStringOleConecction);
                Openconectar();
                TabladeCondicionOferta3 objeto;
                OleDbDataReader reader = TableFacturaMicrosoftAccessDatabaseOleDbCommand.ExecuteReader();
                while (reader.Read())
                {
                    objeto = new TabladeCondicionOferta3();
                    objeto.condOferta = reader[0].ToString();
                    ListaCondicionesOferta3.Add(objeto);
                }
                reader.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return ListaCondicionesOferta3;
            }
            finally
            {
                CerrarConexion();
            }
            return ListaCondicionesOferta3;
        }

      
        
        
        
        
        
        
        
        
        
        
        
        
        
        
  
        
        ///////////////////////////////////----- FUNCION LLENAR DATOS -----///////////////////////////////////////////////////////////////////////////////////////////

        public void LlenarDatos()
        {
                     
           //--------------------------- TITULOS DE CONCEPTO PARA LOS 3 COMBOBOX ----------------------------------------
            Openconectar();
            ListaTituloConcepto = PopulatedAddComboboxConceptoTitulo();
            cbTituConcep1.DisplayMember = "tituConcept";
            cbTituConcep1.ValueMember = "tituConcept";

            Openconectar();
            ListaTituloConcepto2 = PopulatedAddComboboxConceptoTitulo2();
            cbTituConcep2.DisplayMember = "tituConcept";
            cbTituConcep2.ValueMember = "tituConcept";

            Openconectar();
            ListaTituloConcepto3= PopulatedAddComboboxConceptoTitulo3();
            cbTituConcep03.DisplayMember = "tituConcept";
            cbTituConcep03.ValueMember = "tituConcept";

            Openconectar();
            ListaDescripcionTitulo1 = PopulatedAddComboboxDescripcionTitulo1();
            cbDescripConcep1.DisplayMember = "descrConcept";
            cbDescripConcep1.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo2 = PopulatedAddComboboxDescripcionTitulo2();
            cbDescripConcep2.DisplayMember = "descrConcept";
            cbDescripConcep2.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo3 = PopulatedAddComboboxDescripcionTitulo3();
            cbDescripConcep3.DisplayMember = "descrConcept";
            cbDescripConcep3.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo4 = PopulatedAddComboboxDescripcionTitulo4();
            cbDescripConcep4.DisplayMember = "descrConcept";
            cbDescripConcep4.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo5 = PopulatedAddComboboxDescripcionTitulo5();
            cbDescripConcep5.DisplayMember = "descrConcept";
            cbDescripConcep5.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo6 = PopulatedAddComboboxDescripcionTitulo6();
            cbDescripConcep6.DisplayMember = "descrConcept";
            cbDescripConcep6.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo7 = PopulatedAddComboboxDescripcionTitulo7();
            cbDescripConcep7.DisplayMember = "descrConcept";
            cbDescripConcep7.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo8 = PopulatedAddComboboxDescripcionTitulo8();
            cbDescripConcep8.DisplayMember = "descrConcept";
            cbDescripConcep9.ValueMember = "descrConcept";

            Openconectar();
            ListaDescripcionTitulo9 = PopulatedAddComboboxDescripcionTitulo9();
            cbDescripConcep9.DisplayMember = "descrConcept";
            cbDescripConcep9.ValueMember = "descrConcept";


            ////////////////////////---------------------------------- LLENADO DE TABLA DE CONDICIONES --------------------////////////////////////

            Openconectar();
            ListaCondiciones1 = PopulatedAddComboboxCondiciones1();
            cbCondiciones1.DisplayMember = "condici";
            cbCondiciones1.ValueMember = "condici";


            Openconectar();
            ListaCondiciones2 = PopulatedAddComboboxCondiciones2();
            cbCondiciones2.DisplayMember = "condici";
            cbCondiciones2.ValueMember = "condici";


            Openconectar();
            ListaCondiciones3 = PopulatedAddComboboxCondiciones3();
            cbCondiciones3.DisplayMember = "condici";
            cbCondiciones3.ValueMember = "condici";


            Openconectar();
            ListaCondiciones4 = PopulatedAddComboboxCondiciones4();
            cbCondiciones4.DisplayMember = "condici";
            cbCondiciones4.ValueMember = "condici";


            ///////////////////////////-----------------------------   LLENADO DE TABLA DE CONDICIONES - OFERTA   --------------------------////////

            Openconectar();
            ListaCondicionesOferta1 = PopulatedAddComboboxCondicion_Oferta1();
            cbCondiOferta1.DisplayMember = "condOferta";
            cbCondiOferta1.ValueMember = "condOferta";

            Openconectar();
            ListaCondicionesOferta2 = PopulatedAddComboboxCondicion_Oferta2();
            cbCondiOferta2.DisplayMember = "condOferta";
            cbCondiOferta2.ValueMember = "condOferta";

            Openconectar();
            ListaCondicionesOferta3 = PopulatedAddComboboxCondicion_Oferta3();
            cbCondiOferta3.DisplayMember = "condOferta";
            cbCondiOferta3.ValueMember = "condOferta";

            //----------------------------------------------------------------------------------------------------

        }

        /////////////////////////////////////// ------------------ INICIAR DATOS ----------- ////////////////////////////////////////////////////

        public void IniciarDatos()
        {
            Openconectar();
            ListaDatosOferta = PopulateDataGridViewFromMicrosoftAccessDatabase();            
            dataGridView1.DataSource = ListaDatosOferta;
            ////--- LLENAR AL COMBOBOX DATOS DE CONCEPTO DE TITULO ----- /////
            LlenarDatos();


            cbTituConcep1.DataSource = ListaTituloConcepto;
            cbTituConcep2.DataSource = ListaTituloConcepto2;
            cbTituConcep03.DataSource = ListaTituloConcepto3;
            cbDescripConcep1.DataSource = ListaDescripcionTitulo1;
            cbDescripConcep2.DataSource = ListaDescripcionTitulo2;
            cbDescripConcep3.DataSource = ListaDescripcionTitulo3;
            cbDescripConcep4.DataSource = ListaDescripcionTitulo4;
            cbDescripConcep5.DataSource = ListaDescripcionTitulo5;
            cbDescripConcep6.DataSource = ListaDescripcionTitulo6;
            cbDescripConcep7.DataSource = ListaDescripcionTitulo7;
            cbDescripConcep8.DataSource = ListaDescripcionTitulo8;
            cbDescripConcep9.DataSource = ListaDescripcionTitulo9;
            cbCondiciones1.DataSource = ListaCondiciones1;
            cbCondiciones2.DataSource = ListaCondiciones2;
            cbCondiciones3.DataSource = ListaCondiciones3;
            cbCondiciones4.DataSource = ListaCondiciones4;
            cbCondiOferta1.DataSource = ListaCondicionesOferta1;
            cbCondiOferta2.DataSource = ListaCondicionesOferta2;
            cbCondiOferta3.DataSource = ListaCondicionesOferta3;

        }

        ///////////////////////////////////////   ------------ FUNCION CARGAR DATOS ------------  ///////////////////////////////////////////////////////////////////////////////77

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

        
        //////////////////////////// ---------- CARGA DE DATOS --------------------- /////////////////////////////
    
        private void FormOferta_Load(object sender, EventArgs e)
        {
            CargaInicio();
            cbTituConcep1.Text = string.Empty;
            cbTituConcep2.Text = string.Empty;
            cbTituConcep03.Text = string.Empty;

            cbDescripConcep1.Text = string.Empty;
            cbDescripConcep2.Text = string.Empty;
            cbDescripConcep3.Text = string.Empty;
            cbDescripConcep4.Text = string.Empty;
            cbDescripConcep5.Text = string.Empty;
            cbDescripConcep6.Text = string.Empty;
            cbDescripConcep7.Text = string.Empty;
            cbDescripConcep8.Text = string.Empty;
            cbDescripConcep9.Text = string.Empty;

            cbCondiOferta1.Text = string.Empty;
            cbCondiOferta2.Text = string.Empty;
            cbCondiOferta3.Text = string.Empty;

            cbCondiciones1.Text = string.Empty;
            cbCondiciones2.Text = string.Empty;
            cbCondiciones3.Text = string.Empty;
            cbCondiciones4.Text = string.Empty;

        }



        //////////////////////////////////////////////////// -------------------- INSERTAR DATOS AL DATAGRID -------------------//////////////////////////////////////////
    

        private void btnAdd_Click(object sender, EventArgs e)
        {
            Openconectar();
            OleDbCommand insertDataIntoMSAccessDataBaseOleDbCommand = new OleDbCommand(InsertDataIntoMicrosoftAccessDatabase, AccessDatabaseConecctionStringOleConecction);

            try
            {
                //los nombres de los campos debe ser correspondientes con la base de datos Access
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("codOferta", OleDbType.VarChar).Value = txtCodOferta.Text;
                
                DateTime FechaOferta;
                var FechaOfertaNull = (object)DBNull.Value;
                if (txtFecha.Text != "")
                {
                    if (DateTime.TryParse(txtFecha.Text,out FechaOferta))
                    {
                        FechaOfertaNull = FechaOferta.ToString("dd/MM/yyyy");
                    }
                }
                
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("fecha", OleDbType.VarChar).Value = FechaOfertaNull;
                
                ///////-----------------------------------------------------------------------------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("anho", OleDbType.VarChar).Value = txtanho.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("mes", OleDbType.VarChar).Value = txtMes.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("dia", OleDbType.VarChar).Value = txtDia.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomContac", OleDbType.VarChar).Value = txtSeñor.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("telefono", OleDbType.VarChar).Value = txtTelefono.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("emailClient", OleDbType.VarChar).Value = txtemail.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("nomEmpre", OleDbType.VarChar).Value = txtorganizacion.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtproyecto.Text;

                //-----------------------------------------------------    COMBOBOX       -----------------------------------------------------------------
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("titConcep1", OleDbType.VarChar).Value = cbTituConcep1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("titConcep2", OleDbType.VarChar).Value = cbTituConcep2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("titConcep3", OleDbType.VarChar).Value = cbTituConcep03.Text;

                //-----------------------------------------------------   CANTIDAD      --------------------------------------------------------------------
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant1", OleDbType.VarChar).Value = txtCant1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant2", OleDbType.VarChar).Value = txtCant2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3", OleDbType.VarChar).Value = txtCant3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant2a", OleDbType.VarChar).Value = txtCant2_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3a", OleDbType.VarChar).Value = txtCant3_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3b", OleDbType.VarChar).Value = txtCant3_2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3c", OleDbType.VarChar).Value = txtCant3_3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3d", OleDbType.VarChar).Value = txtCant3_4.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("cant3e", OleDbType.VarChar).Value = txtCant3_5.Text;

                //----------------------------------------------------     UNIDADES         --------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und1", OleDbType.VarChar).Value = cbUnid1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und2", OleDbType.VarChar).Value = cbUnid2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und2a", OleDbType.VarChar).Value = cbUnid2_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3", OleDbType.VarChar).Value = cbUnid3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3a", OleDbType.VarChar).Value = cbUnid3_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3b", OleDbType.VarChar).Value = cbUnid3_2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3c", OleDbType.VarChar).Value = cbUnid3_3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3d", OleDbType.VarChar).Value = cbUnid3_4.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("und3e", OleDbType.VarChar).Value = cbUnid3_5.Text;


                //---------------------------------------------  DESCRIPCION DEL CONCEPTO   ----------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep1", OleDbType.VarChar).Value = cbDescripConcep1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep2", OleDbType.VarChar).Value = cbDescripConcep2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep2a", OleDbType.VarChar).Value = cbDescripConcep3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3", OleDbType.VarChar).Value = cbDescripConcep4.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3a", OleDbType.VarChar).Value = cbDescripConcep5.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3b", OleDbType.VarChar).Value = cbDescripConcep6.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3c", OleDbType.VarChar).Value = cbDescripConcep7.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3d", OleDbType.VarChar).Value = cbDescripConcep8.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("descripConcep3e", OleDbType.VarChar).Value = cbDescripConcep9.Text;


                //-------------------------------------------------     IMPORTE       ---------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import1", OleDbType.VarChar).Value = txtImport1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import2", OleDbType.VarChar).Value = txtImport2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import2a", OleDbType.VarChar).Value = txtImport2_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3", OleDbType.VarChar).Value = txtImport3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3a", OleDbType.VarChar).Value = txtImport3_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3b", OleDbType.VarChar).Value = txtImport3_2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3c", OleDbType.VarChar).Value = txtImport3_3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3d", OleDbType.VarChar).Value = txtImport3_4.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("import3e", OleDbType.VarChar).Value = txtImport3_5.Text;

                //--------------------------------------------------   CONDICIONES - OFERTA      -----------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condOfert1", OleDbType.VarChar).Value = cbCondiOferta1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condOfert2", OleDbType.VarChar).Value = cbCondiOferta2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condOfert3", OleDbType.VarChar).Value = cbCondiOferta3.Text;

                //----------------------------------------------------  CONDICIONES       --------------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condicion0", OleDbType.VarChar).Value = cbCondiciones1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condicion1", OleDbType.VarChar).Value = cbCondiciones2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condicion2", OleDbType.VarChar).Value = cbCondiciones3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("condicion3", OleDbType.VarChar).Value = cbCondiciones4.Text;


                //-------------------------------------------------         TOTALES             -------------------------------------------------------------

                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total1", OleDbType.VarChar).Value = txtTotal1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total2", OleDbType.VarChar).Value = txtTotal2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total3", OleDbType.VarChar).Value = txtTotal2_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total4", OleDbType.VarChar).Value = txtTotal3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total5", OleDbType.VarChar).Value = txtTotal3_1.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total6", OleDbType.VarChar).Value = txtTotal3_2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total7", OleDbType.VarChar).Value = txtTotal3_2.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total8", OleDbType.VarChar).Value = txtTotal3_3.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total9", OleDbType.VarChar).Value = txtTotal3_4.Text;
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Total1", OleDbType.VarChar).Value = txtTotal3_5.Text;


                //-------------------------------------------------      SUBTOTAL + IGV + TOTAL     -----------------------------------------------------------
                ///----  SUBTOTAL -----/////

               
               insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("Subtotal", OleDbType.VarChar).Value = txtSubtotal.Text;
                
              
                /// --- IGV 18 ---- //
               
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("igv18", OleDbType.VarChar).Value = txtigv18.Text;

                
                insertDataIntoMSAccessDataBaseOleDbCommand.Parameters.AddWithValue("SumaTotal", OleDbType.VarChar).Value = txtTotal.Text;

                
                int DataInsert = insertDataIntoMSAccessDataBaseOleDbCommand.ExecuteNonQuery();
            
                if (DataInsert > 0)
                {
                    //MessageBox.Show("Registro Exitosoo ¡¡¡");
                    MessageBox.Show("Registro Exitoso", "Registro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Refrescar el Datagrid al insertar un registro
                    ListaDatosOferta = PopulateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = ListaDatosOferta;
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
                CerrarConexion();
            }
        }



        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtID.Text))
            {
                MessageBox.Show("Seleccione el Registro que desee Actualizar");
            }
            else
            {
                try
                {
                    Openconectar();
                    OleDbCommand updateDatabaseAccessOleDbCommand = new OleDbCommand(UpdateMicrosoftAccessDataQuery, AccessDatabaseConecctionStringOleConecction);


                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("codOferta", OleDbType.VarChar).Value = txtCodOferta.Text;

                    DateTime fechax;
                    var FechaOfertaNull = (object)DBNull.Value;
                    if (txtFecha.Text != "")
                    {
                        if (DateTime.TryParse(txtFecha.Text, out fechax))
                        {
                            FechaOfertaNull = fechax.ToString("dd/MM/yyyy");
                        }
                    }
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("fecha", OleDbType.VarChar).Value = FechaOfertaNull;

                    ///////-----------------------------------------------------------------------------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("anho", OleDbType.VarChar).Value = txtanho.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("mes", OleDbType.VarChar).Value = txtMes.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("dia", OleDbType.VarChar).Value = txtDia.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("nomContac", OleDbType.VarChar).Value = txtSeñor.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("telefono", OleDbType.VarChar).Value = txtTelefono.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("emailClient", OleDbType.VarChar).Value = txtemail.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("nomEmpre", OleDbType.VarChar).Value = txtorganizacion.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("proyecto", OleDbType.VarChar).Value = txtproyecto.Text;

                    //-----------------------------------------------------    COMBOBOX       -----------------------------------------------------------------
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("titConcep1", OleDbType.VarChar).Value = cbTituConcep1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("titConcep2", OleDbType.VarChar).Value = cbTituConcep2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("titConcep3", OleDbType.VarChar).Value = cbTituConcep03.Text;

                    //-----------------------------------------------------   CANTIDAD      --------------------------------------------------------------------
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant1", OleDbType.VarChar).Value = txtCant1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant2", OleDbType.VarChar).Value = txtCant2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3", OleDbType.VarChar).Value = txtCant3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant2a", OleDbType.VarChar).Value = txtCant2_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3a", OleDbType.VarChar).Value = txtCant3_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3b", OleDbType.VarChar).Value = txtCant3_2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3c", OleDbType.VarChar).Value = txtCant3_3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3d", OleDbType.VarChar).Value = txtCant3_4.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("cant3e", OleDbType.VarChar).Value = txtCant3_5.Text;

                    //----------------------------------------------------     UNIDADES         --------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und1", OleDbType.VarChar).Value = cbUnid1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und2", OleDbType.VarChar).Value = cbUnid2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und2a", OleDbType.VarChar).Value = cbUnid2_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3", OleDbType.VarChar).Value = cbUnid3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3a", OleDbType.VarChar).Value = cbUnid3_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3b", OleDbType.VarChar).Value = cbUnid3_2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3c", OleDbType.VarChar).Value = cbUnid3_3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3d", OleDbType.VarChar).Value = cbUnid3_4.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("und3e", OleDbType.VarChar).Value = cbUnid3_5.Text;


                    //---------------------------------------------  DESCRIPCION DEL CONCEPTO   ----------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep1", OleDbType.VarChar).Value = cbDescripConcep1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep2", OleDbType.VarChar).Value = cbDescripConcep2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep2a", OleDbType.VarChar).Value = cbDescripConcep3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3", OleDbType.VarChar).Value = cbDescripConcep4.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3a", OleDbType.VarChar).Value = cbDescripConcep5.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3b", OleDbType.VarChar).Value = cbDescripConcep6.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3c", OleDbType.VarChar).Value = cbDescripConcep7.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3d", OleDbType.VarChar).Value = cbDescripConcep8.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("descripConcep3e", OleDbType.VarChar).Value = cbDescripConcep9.Text;


                    //-------------------------------------------------     IMPORTE       ---------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import1", OleDbType.VarChar).Value = txtImport1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import2", OleDbType.VarChar).Value = txtImport2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import2a", OleDbType.VarChar).Value = txtImport2_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3", OleDbType.VarChar).Value = txtImport3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3a", OleDbType.VarChar).Value = txtImport3_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3b", OleDbType.VarChar).Value = txtImport3_2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3c", OleDbType.VarChar).Value = txtImport3_3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3d", OleDbType.VarChar).Value = txtImport3_4.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("import3e", OleDbType.VarChar).Value = txtImport3_5.Text;

                    //--------------------------------------------------   CONDICIONES - OFERTA      -----------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condOfert1", OleDbType.VarChar).Value = cbCondiOferta1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condOfert2", OleDbType.VarChar).Value = cbCondiOferta2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condOfert3", OleDbType.VarChar).Value = cbCondiOferta3.Text;

                    //----------------------------------------------------  CONDICIONES       --------------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condicion0", OleDbType.VarChar).Value = cbCondiciones1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condicion1", OleDbType.VarChar).Value = cbCondiciones2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condicion2", OleDbType.VarChar).Value = cbCondiciones3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("condicion3", OleDbType.VarChar).Value = cbCondiciones4.Text;


                    //-------------------------------------------------         TOTALES             -------------------------------------------------------------

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total1", OleDbType.VarChar).Value = txtTotal1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total2", OleDbType.VarChar).Value = txtTotal2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total3", OleDbType.VarChar).Value = txtTotal2_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total4", OleDbType.VarChar).Value = txtTotal3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total5", OleDbType.VarChar).Value = txtTotal3_1.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total6", OleDbType.VarChar).Value = txtTotal3_2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total7", OleDbType.VarChar).Value = txtTotal3_2.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total8", OleDbType.VarChar).Value = txtTotal3_3.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total9", OleDbType.VarChar).Value = txtTotal3_4.Text;
                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Total1", OleDbType.VarChar).Value = txtTotal3_5.Text;


                    //-------------------------------------------------      SUBTOTAL + IGV + TOTAL     -----------------------------------------------------------
                    ///----  SUBTOTAL -----/////


                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("Subtotal", OleDbType.VarChar).Value = txtSubtotal.Text;

                    /// --- IGV 18 ---- //

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("igv18", OleDbType.VarChar).Value = txtigv18.Text;

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("SumaTotal", OleDbType.VarChar).Value = txtTotal.Text;

                    updateDatabaseAccessOleDbCommand.Parameters.AddWithValue("idoferta", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);

                    int DataUpdate = updateDatabaseAccessOleDbCommand.ExecuteNonQuery();

                    if (DataUpdate > 0)
                    {
                        MessageBox.Show("Actualización Exitoso", "Actualizacion de Datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    ListaDatosOferta = PopulateDataGridViewFromMicrosoftAccessDatabase();
                    dataGridView1.DataSource = ListaDatosOferta;
                    CerrarConexion();
                }

            }
        }

        

        private void btnCalcular1_Click(object sender, EventArgs e)
        {
            double Calcular1;
            Calcular1 = Convert.ToDouble(txtCant1.Text) * Convert.ToDouble(txtImport1.Text);
            txtTotal1.Text = Convert.ToString(Calcular1);
        }

        private void btnCalcular2_Click(object sender, EventArgs e)
        {

            double Calcular2;
            Calcular2 = Convert.ToDouble(txtCant2.Text) * Convert.ToDouble(txtImport2.Text);
            txtTotal2 .Text = Convert.ToString(Calcular2);
        }

        private void btnCalcular2_1_Click(object sender, EventArgs e)
        {
            double Calcular2_1;
            Calcular2_1 = Convert.ToDouble(txtCant2_1.Text) * Convert.ToDouble(txtImport2_1.Text);
            txtTotal2_1.Text = Convert.ToString(Calcular2_1);
        }

        private void btnCalcular3_Click(object sender, EventArgs e)
        {
            double Calcular3;
            Calcular3 = Convert.ToDouble(txtCant3.Text) * Convert.ToDouble(txtImport3.Text);
            txtTotal3.Text = Convert.ToString(Calcular3);
        }

        private void btnCalcular3_1_Click(object sender, EventArgs e)
        {
            double Calcular3_1;
            Calcular3_1 = Convert.ToDouble(txtCant3_1.Text) * Convert.ToDouble(txtImport3_1.Text);
            txtTotal3_1.Text = Convert.ToString(Calcular3_1);
        }

        private void btnCalcular3_2_Click(object sender, EventArgs e)
        {
            double Calcular3_2;
            Calcular3_2 = Convert.ToDouble(txtCant3_2.Text) * Convert.ToDouble(txtImport3_2.Text);
            txtTotal3_2.Text = Convert.ToString(Calcular3_2);
        }

        private void btnCalcular3_3_Click(object sender, EventArgs e)
        {
            double Calcular3_3;
            Calcular3_3 = Convert.ToDouble(txtCant3_3.Text) * Convert.ToDouble(txtImport3_3.Text);
            txtTotal3_3.Text = Convert.ToString(Calcular3_3);
        }

        private void btnCalcular3_4_Click(object sender, EventArgs e)
        {
            double Calcular3_4;
            Calcular3_4 = Convert.ToDouble(txtCant3_4.Text) * Convert.ToDouble(txtImport3_4.Text);
            txtTotal3_4.Text = Convert.ToString(Calcular3_4);
        }

        private void btnCalcular3_5_Click(object sender, EventArgs e)
        {
            double Calcular3_5;
            Calcular3_5 = Convert.ToDouble(txtCant3_5.Text) * Convert.ToDouble(txtImport3_5.Text);
            txtTotal3_5.Text = Convert.ToString(Calcular3_5);
        }

        private void btnCalcular_Click(object sender, EventArgs e)
        {

            double CalculoSubtotal;
            CalculoSubtotal = Convert.ToDouble(txtTotal1.Text) + Convert.ToDouble(txtTotal2.Text) + Convert.ToDouble(txtTotal2_1.Text);
            txtSubtotal.Text = Convert.ToString(CalculoSubtotal);
            
           

            double CalculoIGV;
            CalculoIGV = (Convert.ToDouble(txtSubtotal.Text)) * 0.18;
            txtigv18.Text = Convert.ToString(CalculoIGV);
           
            

            double TotalGeneral;
            TotalGeneral = Convert.ToDouble(txtSubtotal.Text) + Convert.ToDouble(txtigv18.Text);
            txtTotal.Text = Convert.ToString(TotalGeneral);
            
    }

        private void button1_Click(object sender, EventArgs e)
        {
            txtCodOferta.Text = txtanho.Text + txtMes.Text + txtDia.Text + txtNumero.Text + txtReferencia.Text.ToUpper() + txtEmpresa.Text.ToUpper();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void limpiarcampos()
        {
            txtCodOferta.Text = string.Empty;
            txtFecha.Text = string.Empty;
            txtSeñor.Text = string.Empty;
            txtTelefono.Text = string.Empty;
            txtemail.Text = string.Empty;
            txtorganizacion.Text = string.Empty;
            txtproyecto.Text = string.Empty;
            txtanho.Text = string.Empty;
            txtMes.Text = string.Empty;
            txtDia.Text = string.Empty;
            txtReferencia.Text = string.Empty;
            txtEmpresa.Text = string.Empty;
            txtNumero.Text = string.Empty;

            cbTituConcep1.Text = string.Empty;
            cbTituConcep2.Text = string.Empty;
            cbTituConcep03.Text = string.Empty;

            cbDescripConcep1.Text = string.Empty;
            cbDescripConcep2.Text = string.Empty;
            cbDescripConcep3.Text = string.Empty;
            cbDescripConcep4.Text = string.Empty;
            cbDescripConcep5.Text = string.Empty;
            cbDescripConcep6.Text = string.Empty;
            cbDescripConcep7.Text = string.Empty;
            cbDescripConcep8.Text = string.Empty;
            cbDescripConcep9.Text = string.Empty;

            txtCant1.Text = string.Empty;
            txtCant2.Text = string.Empty;
            txtCant2_1.Text = string.Empty;
            txtCant3.Text = string.Empty;
            txtCant3_1.Text = string.Empty;
            txtCant3_2.Text = string.Empty;
            txtCant3_3.Text = string.Empty;
            txtCant3_4.Text = string.Empty;
            txtCant3_5.Text = string.Empty;

            cbUnid1.Text = string.Empty;
            cbUnid2.Text = string.Empty;
            cbUnid2_1.Text = string.Empty;
            cbUnid3.Text = string.Empty;
            cbUnid3_1.Text = string.Empty;
            cbUnid3_2.Text = string.Empty;
            cbUnid3_3.Text = string.Empty;
            cbUnid3_4.Text = string.Empty;
            cbUnid3_5.Text = string.Empty;

            cbCondiciones1.Text = string.Empty;
            cbCondiciones2.Text = string.Empty;
            cbCondiciones3.Text = string.Empty;
            cbCondiciones4.Text = string.Empty;

            cbCondiOferta1.Text = string.Empty;
            cbCondiOferta2.Text = string.Empty;
            cbCondiOferta3.Text = string.Empty;

            txtImport1.Text = string.Empty;
            txtImport2.Text = string.Empty;
            txtImport2_1.Text = string.Empty;
            txtImport3.Text = string.Empty;
            txtImport3_1.Text = string.Empty;
            txtImport3_2.Text = string.Empty;
            txtImport3_3.Text = string.Empty;
            txtImport3_4.Text = string.Empty;
            txtImport3_5.Text = string.Empty;

            txtTotal1.Text = string.Empty;
            txtTotal2.Text = string.Empty;
            txtTotal2_1.Text = string.Empty;
            txtTotal3.Text = string.Empty;
            txtTotal3_1.Text = string.Empty;
            txtTotal3_2.Text = string.Empty;
            txtTotal3_3.Text = string.Empty;
            txtTotal3_4.Text = string.Empty;
            txtTotal3_5.Text = string.Empty;

            txtSubtotal.Text = string.Empty;
            txtigv18.Text = string.Empty;
            txtTotal.Text = string.Empty;

        }
        private void btnClean_Click(object sender, EventArgs e)
        {
            limpiarcampos();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            limpiarcampos();

            txtID.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtCodOferta.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            
            if (dataGridView1.CurrentRow.Cells[2].Value != null)
            {
                DateTime fechaoferta = (DateTime) dataGridView1.CurrentRow.Cells[2].Value;
                txtFecha.Text = fechaoferta.ToString("dd/MM/yyyy");
            }
          
            txtanho.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            txtMes.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            txtDia.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            txtSeñor.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            txtTelefono.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            txtemail.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            txtorganizacion.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            txtproyecto.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            
            cbTituConcep1.Text = dataGridView1.CurrentRow.Cells[11].Value.ToString();
            cbTituConcep2.Text = dataGridView1.CurrentRow.Cells[12].Value.ToString();


            //txtReferencia.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //txtEmpresa.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            //txtNumero.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            
            cbTituConcep03.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();

            txtCant1.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
            txtCant2.Text = dataGridView1.CurrentRow.Cells[15].Value.ToString();
            txtCant2_1.Text = dataGridView1.CurrentRow.Cells[16].Value.ToString();
            txtCant3.Text = dataGridView1.CurrentRow.Cells[17].Value.ToString();
            txtCant3_1.Text = dataGridView1.CurrentRow.Cells[18].Value.ToString();
            txtCant3_2.Text = dataGridView1.CurrentRow.Cells[19].Value.ToString();
            txtCant3_3.Text = dataGridView1.CurrentRow.Cells[20].Value.ToString();
            txtCant3_4.Text = dataGridView1.CurrentRow.Cells[21].Value.ToString();
            txtCant3_5.Text = dataGridView1.CurrentRow.Cells[22].Value.ToString();

            cbUnid1.Text = dataGridView1.CurrentRow.Cells[23].Value.ToString();
            cbUnid2.Text = dataGridView1.CurrentRow.Cells[24].Value.ToString();
            cbUnid2_1.Text = dataGridView1.CurrentRow.Cells[25].Value.ToString();
            cbUnid3.Text = dataGridView1.CurrentRow.Cells[26].Value.ToString();
            cbUnid3_1.Text = dataGridView1.CurrentRow.Cells[27].Value.ToString();
            cbUnid3_2.Text = dataGridView1.CurrentRow.Cells[28].Value.ToString();
            cbUnid3_3.Text = dataGridView1.CurrentRow.Cells[29].Value.ToString();
            
            
            
            cbUnid3_4.Text = dataGridView1.CurrentRow.Cells[30].Value.ToString();
            cbUnid3_5.Text = dataGridView1.CurrentRow.Cells[31].Value.ToString();
            cbDescripConcep1.Text = dataGridView1.CurrentRow.Cells[32].Value.ToString();
            cbDescripConcep2.Text = dataGridView1.CurrentRow.Cells[33].Value.ToString();
            cbDescripConcep3.Text = dataGridView1.CurrentRow.Cells[34].Value.ToString();
            cbDescripConcep4.Text = dataGridView1.CurrentRow.Cells[35].Value.ToString();
            cbDescripConcep5.Text = dataGridView1.CurrentRow.Cells[36].Value.ToString();
            cbDescripConcep6.Text = dataGridView1.CurrentRow.Cells[37].Value.ToString();
            cbDescripConcep7.Text = dataGridView1.CurrentRow.Cells[38].Value.ToString();
            cbDescripConcep8.Text = dataGridView1.CurrentRow.Cells[39].Value.ToString();
            cbDescripConcep9.Text = dataGridView1.CurrentRow.Cells[40].Value.ToString();


            txtImport1.Text = dataGridView1.CurrentRow.Cells[41].Value.ToString();
            txtImport2.Text = dataGridView1.CurrentRow.Cells[42].Value.ToString();
            txtImport2_1.Text = dataGridView1.CurrentRow.Cells[43].Value.ToString();


            txtImport3.Text = dataGridView1.CurrentRow.Cells[44].Value.ToString();
            txtImport3_1.Text = dataGridView1.CurrentRow.Cells[45].Value.ToString();
            txtImport3_2.Text = dataGridView1.CurrentRow.Cells[46].Value.ToString();
            txtImport3_3.Text = dataGridView1.CurrentRow.Cells[47].Value.ToString();
            txtImport3_4.Text = dataGridView1.CurrentRow.Cells[48].Value.ToString();
            txtImport3_5.Text = dataGridView1.CurrentRow.Cells[49].Value.ToString();


            cbCondiOferta1.Text = dataGridView1.CurrentRow.Cells[50].Value.ToString();
            cbCondiOferta2.Text = dataGridView1.CurrentRow.Cells[51].Value.ToString();
            cbCondiOferta3.Text = dataGridView1.CurrentRow.Cells[52].Value.ToString();
            cbCondiciones1.Text = dataGridView1.CurrentRow.Cells[53].Value.ToString();
            cbCondiciones2.Text = dataGridView1.CurrentRow.Cells[54].Value.ToString();
            cbCondiciones3.Text = dataGridView1.CurrentRow.Cells[55].Value.ToString();
            cbCondiciones4.Text = dataGridView1.CurrentRow.Cells[56].Value.ToString();


            txtTotal1.Text = dataGridView1.CurrentRow.Cells[57].Value.ToString();
            txtTotal2.Text = dataGridView1.CurrentRow.Cells[58].Value.ToString();
            txtTotal2_1.Text = dataGridView1.CurrentRow.Cells[59].Value.ToString();
            txtTotal3.Text = dataGridView1.CurrentRow.Cells[60].Value.ToString();
            txtTotal3_1.Text = dataGridView1.CurrentRow.Cells[61].Value.ToString();
            txtTotal3_2.Text = dataGridView1.CurrentRow.Cells[62].Value.ToString();
            txtTotal3_3.Text = dataGridView1.CurrentRow.Cells[63].Value.ToString();
            txtTotal3_4.Text = dataGridView1.CurrentRow.Cells[64].Value.ToString();
            txtTotal3_5.Text = dataGridView1.CurrentRow.Cells[65].Value.ToString();

            txtSubtotal.Text = dataGridView1.CurrentRow.Cells[66].Value.ToString();
            txtigv18.Text = dataGridView1.CurrentRow.Cells[67].Value.ToString();
            txtTotal.Text = dataGridView1.CurrentRow.Cells[68].Value.ToString();



        }

        private void btnErase_Click(object sender, EventArgs e)
        {
            Openconectar();

            try
            {
                if (String.IsNullOrEmpty(txtID.Text))
                {
                    MessageBox.Show("Campo ID Vacio , seleccionar un fila de registro");
                }
                else
                {
                    OleDbCommand DeleteMicrosoftAccessDataQueryOleDbCommand = new OleDbCommand(DeleteMicrosoftAccessDataQuery, AccessDatabaseConecctionStringOleConecction);
                    DeleteMicrosoftAccessDataQueryOleDbCommand.Parameters.AddWithValue("idoferta", OleDbType.Integer).Value = Convert.ToInt16(txtID.Text);
                    //abrir la coneexion


                    int DeleteMicrosoftAccessData = DeleteMicrosoftAccessDataQueryOleDbCommand.ExecuteNonQuery();
                    if (DeleteMicrosoftAccessData > 0)
                    {
                        MessageBox.Show("Registro Borrado");
                        //Se debe refrescar el datagrid antes de actualizar los datos eliminados de la base de datos de ACCESSS
                        //limpiar los capos
                        limpiarcampos();
                        ListaDatosOferta = PopulateDataGridViewFromMicrosoftAccessDatabase();
                        dataGridView1.DataSource = ListaDatosOferta;
                        
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
                CerrarConexion();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormReporte reporte = new FormReporte();
            reporte.Show();
        }
    }
}
