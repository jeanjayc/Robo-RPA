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
            string CEP = String.Empty;
            //
            string fileName = @"C:\Users\jeanm\source\repos\BOTRPA\Lista_de_CEPs - DESAFIO RPA - Copy.xlsx";
            var xls = new XLWorkbook(fileName);
            var planilha = xls.Worksheets.First(w => w.Name == "Lista de CEPs");
            var totalLinhas = planilha.Rows().Count();
            // primeira linha é o cabecalho

            List<string> CEPs = new List<string>();
            for (int i = 2; i <= totalLinhas; i++)
            {
                CEPs.Add(planilha.Cell($"B{i}").Value.ToString());
            }
            IWebDriver driver = new ChromeDriver(@"C:\WebDriver\bin\WebDriver\bin\");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            var result = new List<string>();
            var result2 = new List<string>();
            int contadorCEP = 0;
            foreach (var cep in CEPs)
            {
                driver.Navigate().GoToUrl("http://www.buscacep.correios.com.br/sistemas/buscacep/");
                driver.FindElement(By.Name("relaxation")).SendKeys(cep + OpenQA.Selenium.Keys.Enter);
                inputCEP.Text = cep;
                List<IWebElement> elements = driver.FindElements(By.TagName("th")).ToList();
                foreach (var element in elements)
                {
                    result.Add(element.Text);
                }
                List<IWebElement> elements1 = driver.FindElements(By.TagName("td")).ToList();
                foreach (var element in elements1)
                {
                    result2.Add(element.Text);
                }
                contadorCEP++;
                qtCEP.Text = contadorCEP.ToString();
            }
            MontaExcel(result, result2);

        }

        private void MontaExcel(List<string> elements, List<string> elements1)
        {

            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("Lista de CEP's");
            string path = @"C:\Users\jeanm\source\repos\BOTRPA\Resultado.xlsx";

            for (int i = 0; i <= 4; i++)
            {
                if(i == 4)
                {
                    worksheet.Cell(1, 5).Value = "Data da Busca";
                    workbook.SaveAs(path);
                    break;
                }
                worksheet.Cell(1, i + 1).Value = elements[i].ToString();
                workbook.SaveAs(path);
            }

            List<string> resultado = new List<string>();

            foreach (var element in elements1)
            {
                resultado.Add(element.ToString());
            }

            int row = 2;
            int column = 1;
            for (int i = 0; i < resultado.Count; i++)
            {
                if (column == 5)
                {
                    worksheet.Cell(row, column).Value = DateTime.Now;
                    column = 1;
                    row++;
                }
                worksheet.Cell(row, column).Value = string.IsNullOrEmpty(resultado[i].ToString()) ? "não informado" : resultado[i].ToString();
                workbook.SaveAs(path);
                column++;
            }
            
        }
    }
}

