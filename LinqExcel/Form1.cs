using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace LinqExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // pega caminho CSV
                string arquivo_CSV = "" + @"C:\Users\LGMONEZI\Desktop\testlink\Cadastro de Proposta.xls";
                XDocument docXML = ConvertExcel_XML.ConversorExcel_XML(arquivo_CSV);
                //Salva o arquivo XML
                docXML.Save(arquivo_CSV + ".xml");
                //escreve a primeira linha do XML no textBox
                textBox2.Text = docXML.Declaration.ToString();
                //escreve cada linha do XML no TextBox
                foreach (XElement c in docXML.Elements())
                {
                    textBox2.Text += c.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    
    }
}
