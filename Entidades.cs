using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace facturacion
{
    public class Entidades
    {
        public class TabFactura
        {

            // DECLARACION DE TIPO DE VARIABLES PARA EL FORMULARIO FACTURA ----------------

            public int Id { get; set; }
            public string nomb_emp { get; set; }
            public string proyecto { get; set; }
            public string tipofactura { get; set; }
            public string numFactura { get; set; }
            public string ruc_cl { get; set; }
            public float import { get; set; }
            public float IGV { get; set; }
            public float conIgv { get; set; }

            public float detraccion12 { get; set; }

            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? fechaCobroDetra { get; set; }
            //public String fechaCobroDetraEnvia { get; set; }

            public String mes_CobDetra { get; set; }
            public String anho_CobDetra { get; set; }

            /*Cobró Neto de la Cuenta (Pagó Final con desc. SUNAT)*/
            public float aCobrarEnCuenta { get; set; }

            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? fechapagoCliente { get; set; }
            //public String fechaCobroEnvia { get; set; }

            public String mes_pagCli { get; set; }
            public String anho_pagCli { get; set; }


            /*Fecha de Entrega de los Trabajos*/
            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? entregaTrabajos { get; set; }
            //public String entregaTrabajosEnvia { get; set; }


            /*Fecha de Aprobación de los Trabajos*/
            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? aprobacionTrabajos { get; set; }
            //public String aprobacionTrabajosEnvia { get; set; }

            /*Fecha de Emisión*/
            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? fechaEmision { get; set; }
            //public String fechaEmisionEnvio { get; set; }
            public String mes_Emi { get; set; }
            public String anho_Emi { get; set; }

            /*Fecha de Vencimiento de Pago*/
            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? fechaVencimientoPago { get; set; }
            //public String fechaVencimientoPagoEnvio { get; set; }

            /*Fecha de Prevista*/
            /*se debe poner 2 fechas : 1 con Datetime y otra de String*/
            public DateTime? fechaPrevista { get; set; }
            //public String fechaPrevistaEnvia { get; set; }  (AGREGAR CUANDO EXISTA PROCEDIMIENTO ALMACENADO)

            public String observacionesPag { get; set; }

            public DateTime? fechaInicioTrabajo { get; set; }
            public DateTime? fechaInicioContrato { get; set; }
            public DateTime? fechaFinalContrato { get; set; }
            public DateTime? fechaRealTrabajo { get; set; }
            public DateTime? fechaFinalPrevista { get; set; }

            ///-----------------------------------------------------------------------------------------------

        }

        // DECLARACION DE VARIABLES PARA FORMULARIO DE DATOS DE CLIENTE -----------------------------------

        public class TabClientes
        {
            public int Id { get; set; }
           // repite en la tabla de factura ruc_cl
            public string ruc_cl { get; set; }
            //repite en la tabla de factura ruc_cl
            public string nomb_emp { get; set; }
            public string direccion { get; set; }
            public string pagWeb { get; set; }
            public string telefonoFijo { get; set;}
            public string correoEmpresa { get; set; }
            public string contacto { get; set; }
            public string cargo { get; set; }
            public string celular { get; set; }
            public string correo { get; set; }

        }

        public class TablaGasto
        {
            public int Id { get; set; }
            public string nomb_emp { get; set; }
            public string nomb_prove { get; set; }
            public string ruc_pr { get; set; }
            public string proyecto { get; set; }
            public string tipo { get; set; }
            public DateTime? fechaRecibo { get; set; }
            public string Mes_Reci { get; set; }
            public string Anho_Reci { get; set; }
            public string numFactClient { get; set; }
            public string ruc_cl { get; set; }
            public float costo { get; set; }
            public float ConIgv { get; set; }
            public string concepto { get; set; }
            public string observaciones { get; set; }
            public string PeriImputaci { get; set; }
            public string quienPag { get; set; }
            public DateTime? fechaPagEmpresa { get; set; }
            public string Mes_Emp { get; set; }
            public string Anho_Emp { get; set; }

            public string NumFactProve { get; set; }
 
        }

        public class TablaConceptoOferta
        {
            public string tituConcept { get; set; }

        }

        public class TablaConceptGasto
        {
            public string ConceptoGasto { get; set; }
        }

        public class Variablesglobales
        {
           public static String rutaAccess = @"";
          ///  public string rutaAccess { get; set; }
            //public static String AccessDatabaseeConecctionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = "; //LEE LA BASE DE DATOS

        }

        
        //////////////////////// ------------ CLASE DE ENTIDAD - TITULO DE CONCEPTO ---------------------------------------/////////////////////7
        public class TablaTituloConceptoOferta
        {
            public string tituConcept { get; set; }
        }


        public class TablaTituloConceptoOferta2
        {
            public string tituConcept { get; set; }
        }


        public class TablaTituloConceptoOferta3
        {
            public string tituConcept { get; set; }
        }

        ///////////////////////////////////////////////////----------  CLASE ENTIDADES PARA DESCRIPCION DEL TITULO ------///////////////////////////////////////////////////////
       public class TablaDescripcionTitulo1
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo2
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo3
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo4
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo5
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo6
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo7
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo8
        {
            public string descrConcept { get; set; }
        }

        public class TablaDescripcionTitulo9
        {
            public string descrConcept { get; set; }
        }


       
        /// /////////////////////////////////////////  --------  TABLA DE OFERTAS --------- /////////////////////////////////////////////////
        

        public class TablaOfertas
        {
            public int idoferta { get; set; }
            public string codOferta { get; set; }
            public DateTime? fecha { get; set; }
            public string anho { get; set; }
            public string mes { get; set; }
            public string dia { get; set; }
            public string nomContac { get; set; }
            public string telefono { get; set; }
            public string emailClient { get; set; }
            public string nomEmpre { get; set;}
            public string proyecto { get; set; }
            public string titConcep1 { get; set; }
            public string titConcep2 { get; set; }
            public string titConcep3 { get; set; }
            public int cant1 { get; set; }
            public int cant2 { get; set; }
            public int cant3 { get; set; }
            public int cant2a { get; set; }
            public int cant3a { get; set; }
            public int cant3b { get; set; }
            public int cant3c { get; set; }
            public int cant3d { get; set; }
            public int cant3e { get; set; }
            public string und1 { get; set; }
            public string und2 { get; set; }
            public string und2a { get; set; }
            public string und3 { get; set; }
            public string und3a { get; set; }
            public string und3b { get; set; }
            public string und3c { get; set; }
            public string und3d { get; set; }
            public string und3e { get; set; }
            public string descripConcep1 { get; set; }
            public string descripConcep2 { get; set; }
            public string descripConcep2a { get; set; }
            public string descripConcep3 { get; set; }
            public string descripConcep3a { get; set; }
            public string descripConcep3b { get; set; }
            public string descripConcep3c { get; set; }
            public string descripConcep3d { get; set; }
            public string descripConcep3e { get; set; }
            public float import1 { get; set; }
            public float import2 { get; set; }
            public float import2a { get; set; }
            public float import3 { get; set; }
            public float import3a { get; set; }
            public float import3b { get; set; }
            public float import3c { get; set; }
            public float import3d { get; set; }
            public float import3e { get; set; }
            public string condOfert1 { get; set; }
            public string condOfert2 { get; set; }
            public string condOfert3 { get; set; }
            public string condicion0 { get; set; }
            public string condicion1 { get; set; }
            public string condicion2 { get; set; }
            public string condicion3 { get; set; }
            public float Total1 { get; set; }
            public float Total2 { get; set; }
            public float Total3 { get; set; }
            public float Total4 { get; set; }
            public float Total5 { get; set; }
            public float Total6 { get; set; }
            public float Total7 { get; set; }
            public float Total8 { get; set; }
            public float Total9 { get; set; }
            public float Subtotal { get; set; }
            public float igv18 { get; set; }
            public float SumaTotal { get; set; }

        }

        public class TabladeCondiciones1
        {
            public string condici { get; set; }
        }


        public class TabladeCondiciones2
        {
            public string condici { get; set; }
        }

        public class TabladeCondiciones3
        {
            public string condici { get; set; }
        }

        public class TabladeCondiciones4
        {
            public string condici { get; set; }
        }

        /////////////////////////////////////  -------------------------  CLASES DE ENTIDADES PARA TABLA DE CONDICIONES - OFERTA -------///////////////
        
        public class TabladeCondicionOferta1
        {
            public string condOferta { get; set; }
        }

        public class TabladeCondicionOferta2
        {
            public string condOferta { get; set; }
        }

        public class TabladeCondicionOferta3
        {
            public string condOferta { get; set; }
            
        }

     


    }
}
