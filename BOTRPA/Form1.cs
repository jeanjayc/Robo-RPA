using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOTRPA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Ler Excel
            string fileName = @"C:\Users\jeanm\OneDrive\Documentos\Desafio RPA Paschoalotto\Arquivo_de_entrada.xlsx";
            var xls = new XLWorkbook(fileName);
            var worksheet = xls.Worksheets.First(w => w.Name == "Lista de CEPs");
            var totalRows = worksheet.Rows().Count();

            //Add a uma lista os CEP's extraidos do excel
            List<string> CEPs = new List<string>();
            for (int i = 2; i <= totalRows; i++)
            {
                CEPs.Add(worksheet.Cell($"B{i}").Value.ToString());
            }
            //Configuracao Selenium Chrome
            IWebDriver driver = new ChromeDriver(@"C:\WebDriver\bin\WebDriver\bin\");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            //lista que ira receber os titulos dos campos a partir do html
            var headerElements = new List<string>();

            //lista que ira receber os resultados da busca de cada CEP a partir do html
            var bodyElements = new List<string>();

            //contador que ira mostra a quantidade de ceps extraidos
            int counterCEP = 0;
            foreach (var cep in CEPs)
            {
                //Abre o navegador
                driver.Navigate().GoToUrl("http://www.buscacep.correios.com.br/sistemas/buscacep/");

                //Seleciona a tag onde deve ser inserido o CEP no site dos correios
                driver.FindElement(By.Name("relaxation")).SendKeys(cep + OpenQA.Selenium.Keys.Enter);
                //Exibe o CEP atual que esta sendo buscado na interface do programa
                inputCEP.Text = cep;

                //buscando os dados dos theader para montar o cabalho do excel
                List<IWebElement> elements = driver.FindElements(By.TagName("th")).ToList();
                //populando o cabecalho

                foreach (var element in elements)
                {
                    headerElements.Add(element.Text);
                }

                //buscando os dados do corpo para montar a planilha do excel com os resultados da busca de cada cep
                List<IWebElement> elements1 = driver.FindElements(By.TagName("td")).ToList();

                //populando a planilha
                foreach (var element in elements1)
                {
                    bodyElements.Add(element.Text);
                }
                counterCEP++;
                qtCEP.Text = counterCEP.ToString();
            }
            MontaExcel(headerElements, bodyElements);

        }

        private void MontaExcel(List<string> headerElements, List<string> bodyElements)
        {

            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Lista de CEP's");
            string path = @"C:\Users\jeanm\OneDrive\Documentos\Desafio RPA Paschoalotto\Resultado.xlsx";

            for (int i = 0; i <= 4; i++)
            {
                if(i == 4)
                {
                    worksheet.Cell(1, 5).Value = "Data da Busca";
                    workbook.SaveAs(path);
                    break;
                }
                worksheet.Cell(1, i + 1).Value = headerElements[i].ToString();
                workbook.SaveAs(path);
            }

            List<string> resultado = new List<string>();

            foreach (var element in bodyElements)
            {
                resultado.Add(element.ToString());
            }

            int row = 2;
            int column = 1;
            for (int i = 0; i < resultado.Count; i++)
            {
                if (column == 5)
                {
                    
                    column = 1;
                    row++;
                }
                worksheet.Cell(row, column).Value = string.IsNullOrEmpty(resultado[i].ToString()) ? "não informado" : resultado[i].ToString();
                worksheet.Cell(row, 5).Value = DateTime.Now.ToString("dd-MMM-yyyy-HH:mm:ss");
                workbook.SaveAs(path);
                column++;
            }
            
        }
    }
}

