using System;
using System.IO;
using System.Drawing;
using NPOI.XSSF.UserModel;
using ColorMine.ColorSpaces;

public class Program
{
    double h, s, v;
    static void Main()
    {
        Inicio();
    }



    static int linhaExcel; //Indica qual linha da planilha receberá os dados
    static string nomefoto = ""; //copia o nome de cada foto

    static string pastadefotos = "C:\\Users\\silas\\Desktop\\digitainova\\"; // caminho da pasta com as fotos originais tiradas das caixas 

    static string fotosdigitalizadas = "C:\\Users\\silas\\Desktop\\digitalizadas\\"; // caminho da pasta onde as fotos pós digitalizadas serão armazenadas

    static string pathExcel = @"C:\Users\silas\Desktop\digitalizadas\EGL.xlsx"; // Caminho da Planilha Excel que será gerada 
    public static void Inicio()
    {


        int linha = 1;
        //carrega todas as imagens de 11 a 194
        for (int i = 1; i < 20; i++)
        {
            for (int j = 1; j < 5; j++)
            {
                if (i < 10)
                {
                    linhaExcel = linha;
                    string path = pastadefotos + "0" + i + "" + j + ".jpg";
                    nomefoto = "0" + i + "" + j;
                    try
                    {
                        Bitmap myBitmap = new Bitmap(path);
                        VerificarPixel(myBitmap);
                        linha++;
                    }
                    catch
                    {

                        VerificarPixelnulo(nomefoto);
                        linha++;
                    }



                }
                else
                {
                    // Console.WriteLine(i + "" + j + ".jpg");
                    linhaExcel = linha;
                    //Console.WriteLine("linha excel" + linha);
                    string path = pastadefotos + i + "" + j + ".jpg";
                    nomefoto = i + "" + j;
                    try
                    {
                        // Console.WriteLine(nomefoto + " existe");

                        Bitmap myBitmap = new Bitmap(path);
                        VerificarPixel(myBitmap);
                        linha++;
                    }
                    catch
                    {
                        // Console.WriteLine(nomefoto + " não existe");

                        VerificarPixelnulo(nomefoto);
                        linha++;
                    }

                }
            }
        }

        // linhaExcel = 1;
        // Bitmap myBitmap1 = new Bitmap("C:\\Temp\\imagenscolorpixel\\064.jpg");
        // Console.WriteLine("start");
        // VerificarPixel(myBitmap1);


        // linhaExcel = 2;
        // Bitmap myBitmap2 = new Bitmap("C:\\Temp\\\\Nova pasta\\192_2.jpg");
        // Console.WriteLine("191 Meio.");
        // VerificarPixel(myBitmap2);

        // Console.WriteLine("192 Concluido, começa 193.");

        // linhaExcel = 3;
        // Bitmap myBitmap3 = new Bitmap("C:\\Temp\\\\Nova pasta\\193_2.jpg");
        // Console.WriteLine("191 Meio.");
        // VerificarPixel(myBitmap3);

        // Console.WriteLine("193 Concluido, começa 194.");

        // linhaExcel = 4;
        // Bitmap myBitmap4 = new Bitmap("C:\\Temp\\\\Nova pasta\\194_2.jpg");
        // VerificarPixel(myBitmap4);

        // Console.WriteLine("194 Concluido.");

    }

    public static void VerificarPixelnulo(string nomefoto)
    {

        int[,] array = new int[1000, 1610];
        Color pixelColor, pixelcolor2;

        const int linhaMax = 1000;
        const int colunaMax = 1610;

        int linha;
        int coluna;

        List<float> cores = new List<float>();

        Bitmap myBitmap = new Bitmap(colunaMax, linhaMax);
        for (linha = 0; linha < linhaMax; linha++)
        {
            for (coluna = 0; coluna < colunaMax; coluna++)
            {
                array[linha, coluna] = 0;
                myBitmap.SetPixel(coluna, linha, Color.Black);
            }
        }
        // Driver Code
        myBitmap.Save(fotosdigitalizadas + nomefoto + ".jpg");
        // Console.WriteLine("Salvar Excel a seguir");


        SalvarExcel(array, linhaMax, colunaMax);
        // Console.WriteLine("Fim da salvar Excel a seguir");
    }
    public static void VerificarPixel(Bitmap myBitmap)
    {
        //int[,] array = new int[624, 1156];
        //Color pixelColor;

        //const int linhaMax = 624;
        //const int colunaMax = 1156;

        //int linha;
        //int coluna;


        int[,] array = new int[1000, 1610];
        Color pixelColor, pixelcolor2;

        const int linhaMax = 1000;
        const int colunaMax = 1610;

        int linha;
        int coluna;

        List<float> cores = new List<float>();

        for (linha = 0; linha < linhaMax; linha++)
        {
            for (coluna = 0; coluna < colunaMax; coluna++)
            {
                pixelColor = myBitmap.GetPixel(coluna, linha);

                if ((pixelColor.G >= 140))
                {
                    int diff = pixelColor.G - pixelColor.B;
                    int diff2 = pixelColor.B - pixelColor.G;
                    if (diff < 16 )
                    {
                        //myBitmap.SetPixel(coluna, linha, Color.Red);
                        array[linha, coluna] = 100;
                        myBitmap.SetPixel(coluna, linha, Color.White);
                    }
                    // else if(diff > -16)
                    // {
                    //     //myBitmap.SetPixel(coluna, linha, Color.Red);
                    //     array[linha, coluna] = 100;
                    //     myBitmap.SetPixel(coluna, linha, Color.White);
                    // }
                    else
                    {
                        array[linha, coluna] = 0;
                        myBitmap.SetPixel(coluna, linha, Color.Black);
                    }

                 
                }
                else
                {
                    array[linha, coluna] = 0;
                    myBitmap.SetPixel(coluna, linha, Color.Black);
                }




            }

        }




        // Driver Code
        myBitmap.Save(fotosdigitalizadas + nomefoto + ".jpg");
        // Console.WriteLine("Salvar Excel a seguir");


        SalvarExcel(array, linhaMax, colunaMax);
        // Console.WriteLine("Fim da salvar Excel a seguir");


    }




    public static void SalvarExcel(int[,] array, int linhaMax, int colunaMax)
    {
        //Nova bilbioteca para manipulação de excel > XSSF
        XSSFWorkbook wb;
        using (var fs = new FileStream(pathExcel, FileMode.Open, FileAccess.Read))
            wb = new XSSFWorkbook(fs);
        var sheet = wb.GetSheet("0101");


        int linha;
        int coluna;

        long soma = 0;



        float media;

        for (coluna = 0; coluna < colunaMax; coluna++)
        {
            for (linha = 0; linha < linhaMax; linha++)
            {
                soma = array[linha, coluna] + soma;
            }


            media = (float)soma / (float)linhaMax;
  


            var row = sheet.GetRow(linhaExcel) ?? sheet.CreateRow(linhaExcel);
            var cell = row.GetCell(coluna + 1) ?? row.CreateCell(coluna + 1);

            cell.SetCellValue(media);
            soma = 0;



        }
        // Console.WriteLine("Write");
        using (var fs = new FileStream(pathExcel, FileMode.OpenOrCreate, FileAccess.Write))
        {
            wb.Write(fs);
        }
        wb.Close();
    }

}


