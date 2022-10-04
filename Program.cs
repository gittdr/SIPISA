using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AppPisa
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program archivo = new Program();
            archivo.Extraer();
        }
        public void Extraer()
        {
            string[] values;
            DataTable tbl = new DataTable();
            DirectoryInfo dir = new DirectoryInfo(@"\\10.223.208.41\Users\Administrator\Documents\PISA");
            //DirectoryInfo dir = new DirectoryInfo(@"C:\Administración\Proyecto PISA\Ordenes");



            FileInfo[] files = dir.GetFiles("*");
            //FileInfo[] files = dir.GetFiles("*.XLS");
            int count = files.Length;
            if (count > 0)
            {
                foreach (var item in files)
                {
                    string sourceFile = @"\\10.223.208.41\Users\Administrator\Documents\PISA\" + item.Name;
                    //string sourceFile = @"C:\Administración\Proyecto PISA\Ordenes\" + item.Name;
                    string[] strAllLines = File.ReadAllLines(sourceFile, Encoding.UTF8);
                    File.WriteAllLines(sourceFile, strAllLines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray());
                    string lnau = item.Name.ToUpper();
                    string lna = lnau.ToLower();
                    string Ai_orden = lna.Replace(".xls", "");


                    string[] lineas1 = File.ReadAllLines(sourceFile, Encoding.UTF8);
                    lineas1 = lineas1.Skip(1).ToArray();
                    foreach (string line in lineas1)
                    {
                        string renglones = line;
                        char delimitador = '\t';
                        string[] valores = renglones.Split(delimitador);
                        int coln = int.Parse(valores[0]);
                        string col1 = coln.ToString();
                        string col2 = valores[1].ToString();
                        string col3 = valores[2].ToString();
                        string col4 = valores[3].ToString();
                        string col5 = valores[4].ToString();
                        string col6 = valores[5].ToString();
                        string col7 = valores[6].ToString();
                        int cols = int.Parse(valores[7]);
                        string col8 = cols.ToString();
                        string col9 = valores[8].ToString();
                        string col10 = valores[9].ToString();
                        string clave = valores[10].ToString();
                        string Av_cmd_code = clave.Replace("'", "");
                        string descrip = valores[11].ToString().Replace("�", "Ñ");
                        string Av_cmd_description = descrip.Replace("\"", "");
                        string Av_countunit = valores[12].ToString();
                        string col14 = valores[13].ToString();
                        string Af_weight = valores[14].ToString();
                        string col16 = valores[15].ToString();
                        string col17 = valores[16].ToString();
                        string Af_count = Math.Floor(Convert.ToDecimal(valores[17])).ToString();
                        //string Af_count = valores[17].ToString();
                        string Av_weightunit = valores[18].ToString();
                        string col20 = valores[19].ToString();
                        string col21 = valores[20].ToString().Replace("�", "Ñ");
                        string col22 = valores[21].ToString();
                        string col23 = valores[22].ToString();
                        string col24 = valores[23].ToString();
                        int colt = int.Parse(valores[24]);
                        string col25 = colt.ToString();
                        string col26 = valores[25].ToString();
                        string col27 = valores[26].ToString();
                        string col28 = valores[27].ToString();
                        string col29 = valores[28].ToString();
                        string col30 = valores[29].ToString();
                        string col31 = valores[30].ToString();
                        string col32 = valores[31].ToString();
                        string col33 = valores[32].ToString();

                        if (Av_cmd_code != "")
                        {

                            InsertMerc(col1,col2,col3,col4,col5,col6,col7,col8,col9,col10,Av_cmd_code, Av_cmd_description, Av_countunit,col14, Af_weight,col16,col17, Af_count, Av_weightunit,col20,col21,col22,col23,col24,col25,col26,col27,col28,col29,col30,col31,col32,col33);

                        }

                    }
                    string destinationFile = @"\\10.223.208.41\Users\Administrator\Documents\PISAPROCESADAS\" + item.Name;
                    //string destinationFile = @"C:\Administración\Proyecto PISA\Procesadas\" + item.Name;
                    System.IO.File.Move(sourceFile, destinationFile);
                }
            }

        }
        public void InsertMerc(string col1, string col2, string col3, string col4, string col5, string col6, string col7, string col8, string col9, string col10, string Av_cmd_code, string Av_cmd_description, string Av_countunit, string col14, string Af_weight, string col16, string col17, string Af_count, string Av_weightunit, string col20, string col21, string col22, string col23, string col24, string col25, string col26, string col27, string col28, string col29, string col30, string col31, string col32, string col33)
        {
            string cadena = @"Data source=172.24.16.112; Initial Catalog=TMWSuite; User ID=sa; Password=tdr9312;Trusted_Connection=false;MultipleActiveResultSets=true";
            //DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(cadena))
            {

                using (SqlCommand selectCommand = new SqlCommand("sp_Insert_Merc_Pisa_JC", connection))
                {

                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandTimeout = 1000;
                    selectCommand.Parameters.AddWithValue("@col1", col1);
                    selectCommand.Parameters.AddWithValue("@col2", col2);
                    selectCommand.Parameters.AddWithValue("@col3", col3);
                    selectCommand.Parameters.AddWithValue("@col4", col4);
                    selectCommand.Parameters.AddWithValue("@col5", col5);
                    selectCommand.Parameters.AddWithValue("@col6", col6);
                    selectCommand.Parameters.AddWithValue("@col7", col7);
                    selectCommand.Parameters.AddWithValue("@col8", col8);
                    selectCommand.Parameters.AddWithValue("@col9", col9);
                    selectCommand.Parameters.AddWithValue("@col10", col10);
                    selectCommand.Parameters.AddWithValue("@Av_cmd_code", Av_cmd_code);
                    selectCommand.Parameters.AddWithValue("@Av_cmd_description", Av_cmd_description);
                    selectCommand.Parameters.AddWithValue("@Av_countunit", Av_countunit);
                    selectCommand.Parameters.AddWithValue("@col14", col14);
                    selectCommand.Parameters.AddWithValue("@Af_weight", Af_weight);
                    selectCommand.Parameters.AddWithValue("@col16", col16);
                    selectCommand.Parameters.AddWithValue("@col17", col17);
                    selectCommand.Parameters.AddWithValue("@Af_count", Af_count);
                    selectCommand.Parameters.AddWithValue("@Av_weightunit", Av_weightunit);
                    selectCommand.Parameters.AddWithValue("@col20", col20);
                    selectCommand.Parameters.AddWithValue("@col21", col21);
                    selectCommand.Parameters.AddWithValue("@col22", col22);
                    selectCommand.Parameters.AddWithValue("@col23", col23);
                    selectCommand.Parameters.AddWithValue("@col24", col24);
                    selectCommand.Parameters.AddWithValue("@col25", col25);
                    selectCommand.Parameters.AddWithValue("@col26", col26);
                    selectCommand.Parameters.AddWithValue("@col27", col27);
                    selectCommand.Parameters.AddWithValue("@col28", col28);
                    selectCommand.Parameters.AddWithValue("@col29", col29);
                    selectCommand.Parameters.AddWithValue("@col30", col30);
                    selectCommand.Parameters.AddWithValue("@col31", col31);
                    selectCommand.Parameters.AddWithValue("@col32", col32);
                    selectCommand.Parameters.AddWithValue("@col33", col33);



                    try
                    {
                        connection.Open();
                        selectCommand.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        string message = ex.Message;
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }

        }
    }
}
