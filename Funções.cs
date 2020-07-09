
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
//using AutoIt;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using System.Diagnostics;
using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.IE;
using System.Windows;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO.Compression;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.IE;
//using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using System.Diagnostics;
using System.Security.Permissions;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Html5;
using OpenQA.Selenium.Interactions.Internal;

namespace GIC4._0
{
    class Funções
    {
        public void FechaIE()
        {
            foreach (Process p in Process.GetProcessesByName("iexplore"))
            {
                p.Kill();
            }
            System.Threading.Thread.Sleep(1000);
        }

        //============================================================================================================================

        public void Print(String Evidencia, String nometeste)
        {
            Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
            Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(bitmap as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
            String Caminho = Evidencia + nometeste + ".jpg";
            bitmap.Save(Caminho, ImageFormat.Jpeg);
        }

        //===============================================================================================================================

        public void FechaExcel()
        {
            foreach (Process p in Process.GetProcessesByName("Excel"))
            {
                p.Kill();
                System.Threading.Thread.Sleep(1000);
            }
        }
        //===============================================================================================================================

        public void FechaDriver()
        {
            foreach (Process p in Process.GetProcessesByName("IEDriverServer"))
            {
                p.Kill();
                System.Threading.Thread.Sleep(1000);
            }
        }
        //=========================================================================================================================

        public String GeraCpf()
        {

            int soma = 0, resto = 0;
            int[] multiplicador1 = new int[9] { 10, 9, 8, 7, 6, 5, 4, 3, 2 };
            int[] multiplicador2 = new int[10] { 11, 10, 9, 8, 7, 6, 5, 4, 3, 2 };

            Random rnd = new Random();
            string semente = rnd.Next(100000000, 999999999).ToString();

            for (int i = 0; i < 9; i++)
                soma += int.Parse(semente[i].ToString()) * multiplicador1[i];

            resto = soma % 11;
            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;

            semente = semente + resto;
            soma = 0;

            for (int i = 0; i < 10; i++)
                soma += int.Parse(semente[i].ToString()) * multiplicador2[i];

            resto = soma % 11;

            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;

            semente = semente + resto;
            return semente;
        }

        //===========================================================================================================================================================

        public String GeraCNPJ()
        {



            Random rnd = new Random();



            int n1 = rnd.Next(1, 9);

            int n2 = rnd.Next(1, 9);

            int n3 = rnd.Next(1, 9);

            int n4 = rnd.Next(1, 9);

            int n5 = rnd.Next(1, 9);

            int n6 = rnd.Next(1, 9);

            int n7 = rnd.Next(1, 9);

            int n8 = rnd.Next(1, 9);

            int n9 = rnd.Next(1, 9);

            int n10 = rnd.Next(1, 9);

            int n11 = rnd.Next(1, 9);

            int n12 = rnd.Next(1, 9);

            int d1 = n12 * 2 + n11 * 3 + n10 * 4 + n9 * 5 + n8 * 6 + n7 * 7 + n6 * 8 + n5 * 9 + n4 * 2 + n3 * 3 + n2 * 4 + n1 * 5;



            d1 = 11 - (d1 % 11);



            if (d1 >= 10)
            {

                d1 = 0;

            }



            int d2 = d1 * 2 + n12 * 3 + n11 * 4 + n10 * 5 + n9 * 6 + n8 * 7 + n7 * 8 + n6 * 9 + n5 * 2 + n4 * 3 + n3 * 4 + n2 * 5 + n1 * 6;

            d2 = 11 - (d2 % 11);



            if (d2 >= 10)
            {

                d2 = 0;

            }

            String teste;

            teste = Convert.ToString(n1) + Convert.ToString(n2) + Convert.ToString(n3) + Convert.ToString(n4) +

            Convert.ToString(n5) + Convert.ToString(n6) + Convert.ToString(n7) + Convert.ToString(n8) +

            Convert.ToString(n9) + Convert.ToString(n10) + Convert.ToString(n11) +

            Convert.ToString(n12) + Convert.ToString(d1) + Convert.ToString(d2);

            return teste;



        }
    }
}