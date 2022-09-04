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

    static string pastadefotos = "C:\\Users\\silas\\Desktop\\digitainova\\"; // caminho da pasta com as fotos tiradas das caixas 

    static string fotosdigitalizadas = "C:\\Users\\silas\\Desktop\\digitalizadas\\"; // caminho da pasta onde as fotos pós digitalizadas serão armazenadas

    static string pathExcel = @"C:\Users\silas\Desktop\digitalizadas\EGL.xlsx";
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

                // Console.WriteLine(pixelColor.GetBrightness()+"  "+pixelColor.GetSaturation());
                // cores.Add(pixelColor.GetHue());

                // var hsvcolor = new Hsv { H = pixelColor.GetHue(), S = pixelColor.GetSaturation(), V = pixelColor.GetBrightness() };

                // var rgbconvert = rgbcolor.To<Hsv>();
                // if ((hsvcolor.H >330 ) || (hsvcolor.H > -330 ) )
                // {
                // hsvcolor.H = 0;
                // }
                // if (hsvcolor.S < 0)
                // {
                //     hsvcolor.S = hsvcolor.S * -1;
                // }
                // if (hsvcolor.V < 0)
                // {
                //     hsvcolor.V = hsvcolor.V * -1;
                // }

                // var hsvcolor2 = new Hsv { H = 100, S = 0.9, V = hsvcolor.V };

                // // var newrgb = new Rgb{ R = pixelColor.R, G = 100, B=pixelColor.B };
                // var newrgb = hsvcolor2.To<Rgb>();

                // Console.WriteLine((newrgb.R)+"|"+ (newrgb.G)+" | " +(newrgb.B));


                // while (newrgb.R > 255 || newrgb.R < 0 ||
                // newrgb.G > 255 || newrgb.G < 0 ||
                // newrgb.B > 255 || newrgb.B < 0)
                // {
                //     if ((int)newrgb.R > 255)
                //     {
                //         newrgb.R = 255;
                //     }
                //     else if ((int)newrgb.R < 0)
                //     {
                //         newrgb.R = 0;
                //     }
                //     if ((int)newrgb.G > 255)
                //     {
                //         newrgb.G = 255;
                //     }
                //     else if ((int)newrgb.G < 0)
                //     {
                //         newrgb.G = 0;
                //     }
                //     if ((int)newrgb.B > 255)
                //     {
                //         newrgb.B = 255;
                //     }
                //     else if ((int)newrgb.B < 0)
                //     {
                //         newrgb.B = 0;
                //     }
                // }

                // if (newrgb.R > 255 || newrgb.R < 0 ||
                // newrgb.G > 255 || newrgb.G < 0 ||
                //  newrgb.B > 255 || newrgb.B < 0)
                // {
                // }
                // else
                // {
                // myBitmap.SetPixel(coluna, linha, Color.FromArgb((int)newrgb.R, (int)newrgb.G, (int)newrgb.B));

                // }
                //     if (newrgb.R < 145 && newrgb.G < 145 && newrgb.B < 145)
                //     {
                //         newrgb.R = 0;
                //         newrgb.G = 0;
                //         newrgb.B = 0;
                //     }
                //     else if ((newrgb.R - newrgb.G) > 12 ||
                //    (newrgb.R - newrgb.G) < -12 ||
                //    (newrgb.R - newrgb.B) > 12 ||
                //    (newrgb.R - newrgb.B) < -12 ||
                //    (newrgb.G - newrgb.B) > 12 ||
                //    (newrgb.G - newrgb.B) < -12)
                //     {
                //         newrgb.R = 0;
                //         newrgb.G = 0;
                //         newrgb.B = 0;
                //     }

                // myBitmap.SetPixel(coluna, linha, Color.FromArgb((int)newrgb.R,(int)newrgb.G,(int)newrgb.B));
                // if ((newrgb.R >= 0 && newrgb.R <= 255) && (newrgb.G >= 0 && newrgb.G <= 255) && (newrgb.B >= 0 && newrgb.B <= 255))
                // {
                //     // Console.WriteLine(newrgb.R+" "+ newrgb.G+ " "+ newrgb.B);
                //     myBitmap.SetPixel(coluna, linha, Color.FromArgb((int)newrgb.R, (int)newrgb.G, (int)newrgb.B));

                // }
                // Console.WriteLine(hsvcolor);


                // if (newrgb.R != 255 && newrgb.G != 255 && newrgb.B != 255)
                // {
                //    Console.WriteLine("Coluna: " + coluna + " | Linha: " + linha);
                // }


                // if ((newrgb.R >= 105 && newrgb.R <= 192) && (newrgb.G >= 105 && newrgb.G <= 192) && (newrgb.B >= 105 && newrgb.B <= 192))
                // {
                //     //myBitmap.SetPixel(coluna, linha, Color.Red);
                //     array[linha, coluna] = 100;

                //}
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


        // Console.WriteLine("Salvar Excel iniciado");

        //excelSheet.Cells[1, 1] = 1;

        int linha;
        int coluna;

        long soma = 0;


        //for (linha = 0; linha < linhaMax; linha++)
        //{
        //    for (coluna = 0; coluna < colunaMax; coluna++)
        //    {
        //        //excelSheet.Cells[linha + 1, coluna + 1] = array[linha, coluna];

        //        soma = array[linha, coluna] + soma;

        //    }

        //    excelSheet.Cells[1, linha + 1] = soma / colunaMax;
        //    soma = 0;

        //}

        float media;

        for (coluna = 0; coluna < colunaMax; coluna++)
        {
            for (linha = 0; linha < linhaMax; linha++)
            {
                soma = array[linha, coluna] + soma;
            }


            media = (float)soma / (float)linhaMax;
            // Console.WriteLine($"Media coluna {coluna}: {media}");
            //excelSheet.Cells[1, coluna + 1] = media; //Antigo, só gravava na linha 1



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


