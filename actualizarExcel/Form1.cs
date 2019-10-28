using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace actualizarExcel
{
    public partial class Form1 : Form
    {

        bool sigue = true;
        string salida;
        string url;

        public Form1()
        {
            InitializeComponent();
            this.webBrowser1.ScriptErrorsSuppressed = true;            
            this.webBrowser1.DocumentCompleted += WebBrowser1_DocumentCompleted;
        }


        private void btnProcesar_Click(object sender, EventArgs e)
        {

            salida = "";
            this.sigue = true;
            string path = "archivo.csv";

            this.openFileDialog1.FileName = "";
            this.openFileDialog1.Filter = "Comma Separation Values|*.csv";
            this.openFileDialog1.ShowDialog();
            if (!File.Exists(this.openFileDialog1.FileName))
            {
                MessageBox.Show("error, no existe el archivo");
                return;
            }

            path = this.openFileDialog1.FileName;

            var archivo = File.OpenText(path);
            while (!archivo.EndOfStream)
            {
                var linea = archivo.ReadLine();
                this.url = linea.IndexOf(";") == -1 ? linea.Substring(0) : linea.Substring(0, linea.IndexOf(";") );
                this.sigue = false;
                this.webBrowser1.Navigate(url);
                while (!this.sigue) Application.DoEvents();
                //"https://www.nubbeo.com.ar/sensor-digital-temperatura-ds18b20-cable-sumergible-nubbeo-549563105xJM"
            }
            archivo.Close();    
            path = "archivo1.csv";
            File.WriteAllText(path, this.salida);
            System.Diagnostics.Process.Start(Path.GetFullPath(path));

        }

        private void WebBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

            if (e.Url.AbsoluteUri==this.url)
            {
                try
                {
                    //var a = this.webBrowser1.Document.GetElementById("Price").GetAttribute("Value");
                    var a = this.buscarValor();

                    this.salida += e.Url.AbsoluteUri + ";" + a + ";\n";
                }
                catch
                {
                    this.salida += e.Url.AbsoluteUri + ";" + ";\n";
                }
                this.sigue = true;
            }
        }

           private string buscarValor()
        {
            HtmlElementCollection theElementCollection = this.webBrowser1.Document.GetElementsByTagName("span");
            foreach (HtmlElement curElement in theElementCollection)
            {
                if (curElement.GetAttribute("className").ToString() == "money")
                {
                    //MessageBox.Show(curElement.GetAttribute("InnerText")); // Do something you want
                    return curElement.GetAttribute("InnerText").Trim();
                }
            }
            return ""; 
        }
                     

    }

 }


