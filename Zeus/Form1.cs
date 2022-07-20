using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Zeus
{
    public partial class Form1 : Form
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Dot Tutorials";
        static readonly string sheet = "dados";
        static readonly string SpreadsheetId = "1dA3D78a_z9SVu-OvzTr-LMZKPvnLiZXr_FUFLl7dCJU";
        static SheetsService service;
        string escritorio;
        string atividade;
        string sap;
        string instancia;
        string celular;
        string contato;
        string cliente;
        string observacao;
        string servico;
        string colaborador;
        int a = 0;
        int b = 0;
        int c = 0;
        int d = 0;
        int f = 0;
        int g = 0;
        int h = 0;
        string ATIVIDADE = "ATIVIDADE";
        string CONTA = "SAP TECNICO";
        string INSTANCIA = "INSTANCIA";
        string CLIENTE = "CLIENTE";
        string CONTATO = "CONTATO";
        string CELULAR = "CELULAR";
        string OBSERVAÇÃO = "OBSERVAÇÃO";
        string SERVICO = "SERVIÇO";
        string FUNCIONARIO = "COLABORADOR";
        string hora = DateTime.Now.ToShortTimeString();
        string data = DateTime.Now.ToShortDateString();
        public object Clear { get; private set; }
        public Form1()
        {

            InitializeComponent();
            GoogleCredential credential;
            using (var stream = new FileStream("clientes.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            // Create Google Sheets API service.
            var sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName
            });
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Random aleatorio = new Random();
            int valor = aleatorio.Next(1, 50);

            if (textBox3.Text == "ATIVIDADE" || textBox3.Text == "")
            {
                MessageBox.Show("Favor informar ATIVIDADE");
            }
            else if (textBox4.Text == "SAP TECNICO" || textBox4.Text == "")
            {
                MessageBox.Show("Favor informar SAP TECNICO");
            }
            else if (textBox1.Text == "INSTANCIA" || textBox1.Text == "")
            {
                MessageBox.Show("Favor informar INSTANCIA");
            }
            else if (comboBox1.Text == "")
            {
                MessageBox.Show("Favor informar SERVIÇO ");
            }
            else if (comboBox3.Text == "")
            {
                MessageBox.Show("Favor informar ESCRITORIO ");
            }
            else if (textBox6.Text == "CLIENTE" || textBox6.Text == "")
            {
                MessageBox.Show("Favor informar CLIENTE");
            }
            else if (textBox5.Text == "CONTATO" || textBox5.Text == "")
            {
                MessageBox.Show("Favor informar CONTATO");
            }
            else if (textBox2.Text == "CELULAR" || textBox2.Text == "")
            {
                MessageBox.Show("Favor informar CELULAR");
            }
            else if (comboBox2.Text == "")
            {
                MessageBox.Show("Favor informar COLABORADOR ");
            }
            else if (textBox7.Text == "OBSERVAÇÃO" || textBox7.Text == "")
            {
                MessageBox.Show("Favor informar OBSERVAÇÃO");
            }
            else
            {
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(chromeOptions);
                driver.Navigate().GoToUrl($"https://wom.gvt.com.br:19900/wfm/certificationClean/certificationClean.faces?aid={atividade}&userlogin={sap}&atoken=123");

                try
                {
                    driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr/td/input")).Click();
                    String data1 = driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr[1]/td/label")).Text;
                }
                catch
                {
                }
                try
                {
                    driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr[2]/td/input")).Click();
                    String data2 = driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr[2]/td/label")).Text;
                }
                catch
                {
                }
                try
                {
                    driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr[3]/td/input")).Click();
                    String data3 = driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/table/tbody/tr[3]/td/label")).Text;
                }
                catch
                {
                }
                try
                {
                    driver.FindElement(By.XPath("/html/body/form/div[1]/div[2]/input")).Click();
                    Task.Delay(1000).Wait();
                    driver.FindElement(By.XPath("/html/body/div[2]/div[3]/div/div/div/input[1]")).Click();
                    driver.Quit();
                    MessageBox.Show($"Senha de certificação: {valor}\nInformações salvas com sucesso!");
                    var range = $"{sheet}!A:A";
                    var valueRange = new ValueRange();
                    var oblist = new List<object>() { $"Data e hora: {data} - {hora}|Numero da atividade: {atividade}|Login: {sap}|Instancia: {instancia}|Cliente: {cliente}|Contato: {contato}|Celular: {celular}|Serviço: {servico}|Colaborador: {colaborador}|Escritorio: {escritorio}|Senha: {valor}" };
                    valueRange.Values = new List<IList<object>> { oblist };
                    var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    var appendReponse = appendRequest.Execute();
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox7.Clear();
                    textBox1.Text = INSTANCIA;
                    textBox2.Text = CELULAR;
                    textBox3.Text = ATIVIDADE;
                    textBox4.Text = CONTA;
                    textBox5.Text = CONTATO;
                    textBox6.Text = CLIENTE;
                    textBox7.Text = OBSERVAÇÃO;
                    comboBox1.Text = SERVICO;
                    comboBox2.Text = FUNCIONARIO;
                    comboBox3.Text = escritorio;
                    a = 0;
                    b = 0;
                    c = 0;
                    d = 0;
                    f = 0;
                    g = 0;
                    h = 0;
                }
                catch
                {
                    driver.Quit();
                    var range = $"{sheet}!A:A";
                    var valueRange = new ValueRange();
                    var oblist = new List<object>() { $"Data e hora: {data} - {hora}|Numero da atividade: {atividade}|Login: {sap}|Instancia: {instancia}|Cliente: {cliente}|Contato: {contato}|Celular: {celular}|Serviço: {servico}|Colaborador: {colaborador}|Escritorio: {escritorio}|Senha: {valor}" };
                    valueRange.Values = new List<IList<object>> { oblist };
                    var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                    var appendReponse = appendRequest.Execute();
                    MessageBox.Show($"Senha de certificação: {valor}\nInformações salvas com sucesso!");
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox7.Clear();
                    textBox1.Text = INSTANCIA;
                    textBox2.Text = CELULAR;
                    textBox3.Text = ATIVIDADE;
                    textBox4.Text = CONTA;
                    textBox5.Text = CONTATO;
                    textBox6.Text = CLIENTE;
                    textBox7.Text = OBSERVAÇÃO;
                    comboBox1.Text = SERVICO;
                    comboBox2.Text = FUNCIONARIO;
                    comboBox3.Text = escritorio;
                    a = 0;
                    b = 0;
                    c = 0;
                    d = 0;
                    f = 0;
                    g = 0;
                    h = 0;
                }
            }
        }
        public void textBox3_TextChanged(object sender, EventArgs e)
        {
            atividade = textBox3.Text;
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            sap = textBox4.Text;
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            instancia = textBox1.Text;
        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            cliente = textBox6.Text;
        }
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            contato = textBox5.Text;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            celular = textBox2.Text;
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            observacao = textBox7.Text;
        }
        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (a <= 1)
            {
                a++;
                textBox1.Clear();
            }
        }
        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            if (h <= 1)
            {
                h++;
                textBox2.Clear();
            }
        }
        private void textBox3_MouseClick(object sender, MouseEventArgs e)
        {
            if (b <= 1)
            {
                b++;
                textBox3.Clear();
            }
        }
        private void textBox4_MouseClick(object sender, MouseEventArgs e)
        {

            if (c <= 1)
            {
                c++;
                textBox4.Clear();
            }
        }
        private void textBox5_MouseClick(object sender, MouseEventArgs e)
        {
            if (d <= 1)
            {
                d++;
                textBox5.Clear();
            }
        }
        private void textBox6_MouseClick(object sender, MouseEventArgs e)
        {
            if (f <= 1)
            {
                f++;
                textBox6.Clear();
            }
        }
        private void textBox7_MouseClick(object sender, MouseEventArgs e)
        {
            if (g <= 1)
            {
                g++;
                textBox7.Clear();
            }
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            colaborador = comboBox2.Text;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            servico = comboBox1.Text;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            escritorio = comboBox3.Text;
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

    }

}
