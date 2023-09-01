using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace EjercicioTramas.PL
{
    public class Trama
    {

        public static ML.Result TramaExcel()
        {
            ML.Result result = new ML.Result();
            try
            {
                //Lectura del archivo Excel
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0 XML;Data Source=C:\Users\IsacK\OneDrive\Documents\LecturaExcel.xlsx";
                          
                using (OleDbConnection context = new OleDbConnection(connectionString))
                {
                    //Apertura de conección
                    context.Open();    
                

                //Declaración del nombre de la tabla
                DataTable tableTrama new DataTable();


                //Leer última fila del archivo Excel
                DataRow row = tableTrama.Rows.last;

                String base = row[2].ToString();
                String extendida = row[4].ToString()


                ML.Campo campo = new ML.Campo();

                campo.Pedido = int.Parse(base.subString(0, 10));
                campo.ProdADN = int.Parse(base.subString(11, 5));

                }

            }
            catch (Exception ex)
            {
                result.Correct = false;
                result.Message = ex.Message;
                result.Ex = ex;
            }
            return result;

        }
    }
}
