
using GoogleMapsGeocoding;
using GoogleMapsGeocoding.Common;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
                notifyIcon1.Visible = true;
            }
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }


        //INTERVALO DE TEMPO PARA EXECUÇÃO DA ATUALIZAÇÃO TABELA DE VEÍCULOS COM ID E PLACA
        int horaVeic = 24;
        int minutoVeic = 0;
        int segundoVeic = 00;

        //INTERVALO DE TEMPO PARA ATUALIZAÇÃO DAS POSIÇÕES POR ID
        int horaPos = 0;
        int minutoPos = 15;
        int segundoPos = 0;

        int contErro = 0; //INICIA VARIAVEL QUE CONTA ERROS EM CADA CONEXÃO COM SASCAR

        public Form1()
        {
            InitializeComponent();
        }

        private void ExecutaAtualizacao(String att, String agora)
        {

            Servico1.SasIntegraWSService servico = new Servico1.SasIntegraWSService();

            using (var con = new MySql.Data.MySqlClient.MySqlConnection("Server=host;Database=dbname;Uid=root;Pwd=password;"))
            {
                con.Open();

                using (var cmd = new MySql.Data.MySqlClient.MySqlCommand("SELECT idVeiculoSascar, placa FROM veic_aux WHERE idVeiculoSascar >1", con))
                {
                    List<Carro> idVeiculosParaAtualizar = new List<Carro>();
                    MySqlDataReader result = cmd.ExecuteReader();
                    while (result.Read())
                    {
                        Carro veic_att = new Carro();
                        veic_att.IdVeiculo = (int)result["idVeiculoSascar"];
                        veic_att.Placa = (string)result["placa"];

                        idVeiculosParaAtualizar.Add(veic_att);
                    }
                    for (int i = 0; i < idVeiculosParaAtualizar.Count; i++)
                    {
                        Carro veiculo = new Carro();
                        veiculo.IdVeiculo = idVeiculosParaAtualizar[i].IdVeiculo;
                        veiculo.Placa = idVeiculosParaAtualizar[i].Placa;

                        var pacotePosicao = servico.obterPacotePosicaoHistorico("login", "senha", att, agora, veiculo.IdVeiculo, true);

                        if (pacotePosicao != null)

                        {
                            var ultimoPacote = pacotePosicao.Length - 1;
                            veiculo.Latitude = pacotePosicao[ultimoPacote].latitude;
                            veiculo.Longitude = pacotePosicao[ultimoPacote].longitude;
                            veiculo.DataPacote = pacotePosicao[ultimoPacote].dataPacote;
                            veiculo.CidadeAtual = pacotePosicao[ultimoPacote].cidade;
                            veiculo.UFatual = pacotePosicao[ultimoPacote].uf;

                            using (var connection = new MySql.Data.MySqlClient.MySqlConnection("Server=host;Database=dbname;Uid=root;Pwd=password;"))
                            {
                                connection.Open();

                                try
                                {
                                    using (var command = new MySql.Data.MySqlClient.MySqlCommand(@"UPDATE veic_aux
                                                                                                              SET lat = @lat, 
                                                                                                                  lng = @lng,
                                                                                                                  dataAtualizacaoSascar = @dataPacote,
                                                                                                                  cidade_atual = @cidade_atual,
                                                                                                                  uf_atual = @uf_atual,
                                                                                                            WHERE placa = @va.placa", connection))
                                    {
                                        command.Prepare();
                                        command.Parameters.AddWithValue("@va.placa", veiculo.Placa);
                                        command.Parameters.AddWithValue("@lat", veiculo.Latitude);
                                        command.Parameters.AddWithValue("@lng", veiculo.Longitude);
                                        command.Parameters.AddWithValue("@dataPacote", veiculo.DataPacote);
                                        command.Parameters.AddWithValue("@cidade_atual", veiculo.CidadeAtual);
                                        command.Parameters.AddWithValue("@uf_atual", veiculo.UFatual);

                                        command.ExecuteNonQuery();
                                    }
                                    connection.Close();
                                }
                                catch (MySql.Data.MySqlClient.MySqlException)
                                {
                                    connection.Close();
                                    con.Close();
                                }


                            }
                        }
                    }


                }
                con.Close();
            }
        }
        private void AtualizaPos()
        {

            lblLog.Text = ("Atualizando Posições dos veículos na base de dados.\r\nAguarde...");

            DateTime dia = DateTime.Now; // DATA + HORA ATUAL
            var intervalo = 30; // INTERVALO DE TEMPO EM horas QUE SERÃO PESQUISADOS OS PACOTES DE POSIÇÃO POR VEÍCULO.

            var agora = dia.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
            var hora = dia.ToString("HH':'mm':'ss");
            var att = dia.AddMinutes(-intervalo).ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
            var umAnoAtras = dia.AddYears(-4).ToString("yyyy'-'MM'-'dd HH':'mm':'ss");

            try
            {
                ExecutaAtualizacao(att, agora);
            }
            catch
            {
                //DENTRO DESSE CATCH É ONDE TRATAMOS O ERRO DE CONEXÃO E ENVIAMOS EMAIL PARA O GESTOR.
                contErro++;
                if (contErro <= 3)
                {
                    int segundos = 10 * 1000;
                    Thread.Sleep(segundos);
                    AtualizaPos();
                }
                else
                {
                    using (SmtpClient smtp = new SmtpClient())
                    {
                        using (MailMessage email = new MailMessage())
                        {
                            // Servidor SMTP
                            smtp.Host = "smtp.gmail.com";
                            smtp.UseDefaultCredentials = false;
                            smtp.Credentials = new System.Net.NetworkCredential("email remetente", "senha email remetente");
                            smtp.Port = 587;
                            smtp.EnableSsl = true;

                            //Email (Mensagem)
                            email.From = new MailAddress("email remetente");
                            email.To.Add("email destinatario");

                            email.Subject = "Erro no Integrador";
                            email.IsBodyHtml = true;
                            email.Body = "<h1>Atenção</h1><br><p>Olá gestor, houve alguma falha de comunicação entre servidor - cliente, favor reiniciar o aplicativo</p>";

                            //Enviar email
                            smtp.Send(email);
                        }
                    }

                }
            
        }
        lblLog.Text = ("Atualização de Posições finalizada as " + hora + "\r\nPróxima atualização automática em 15 minutos");
        timerAttPos.Enabled = true;
        }

    private void Pesquisa()
    {

    }

    private void button1_Click(object sender, EventArgs e)
    {
        AtualizaPos();
        AtualizaCliente();
        //salvaLogErro();
    }

    private void AtualizaVeic()
    {
        //lblTimerVeic.Text = horaVeic + ":" + minutoVeic + ":" + segundoVeic;

        lblLog.Text = ("Atualizando tabela de veículos na base de dados...");

        DateTime dia = DateTime.Now; // DATA + HORA ATUAL

        var agora = dia.ToString("HH':'mm':'ss");

        Servico1.SasIntegraWSService servico = new Servico1.SasIntegraWSService();

        List<Carro> listCarros = new List<Carro>();
        var pacoteVeiculos = servico.obterVeiculos("login", "senha", 1000, true, 1, false);
        lblLog.Text = (pacoteVeiculos.Length + " veículos na base Sascar");
        lblLog.Text = ("Atualizando base de dados, aguarde...");

        foreach (var pacote in pacoteVeiculos)
        {
            var plac1 = pacote.placa.Substring(0, 3);
            var plac2 = pacote.placa.Substring(3, 4);
            var placFim = (plac1 + '-' + plac2);
            Carro veiculo = new Carro
            {
                IdVeiculo = pacote.idVeiculo,
                Placa = placFim
            };
            listCarros.Add(veiculo);


            using (var connection = new MySql.Data.MySqlClient.MySqlConnection("Server=host;Database=dbname;Uid=root;Pwd=password;"))
            {
                connection.Open();

                try
                {
                    using (var command = new MySql.Data.MySqlClient.MySqlCommand(@"INSERT INTO veic_aux (idVeiculoSascar, placa) 
                                                                                        VALUES (@idVeiculo, @placa);", connection))
                    {
                        command.Prepare();
                        command.Parameters.AddWithValue("@idVeiculo", veiculo.IdVeiculo);
                        command.Parameters.AddWithValue("@placa", veiculo.Placa);

                        command.ExecuteNonQuery();
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException)
                {
                }

                try
                {
                    using (var command = new MySql.Data.MySqlClient.MySqlCommand(@"UPDATE veic_aux SET idVeiculoSascar = @idVeiculo 
                                                                                    WHERE placa = @placa;", connection))
                    {
                        command.Prepare();
                        command.Parameters.AddWithValue("@idVeiculo", veiculo.IdVeiculo);
                        command.Parameters.AddWithValue("@placa", veiculo.Placa);

                        command.ExecuteNonQuery();
                    }
                }
                catch (MySql.Data.MySqlClient.MySqlException)
                {
                }

            }

        }
        lblLog.Text = ("Atualização de Veículos finalizada as " + agora);
        timerAttVeic.Enabled = true;
    }
    public void button2_Click(object sender, EventArgs e)
    {
        AtualizaVeic();
    }

    private void AtualizaCliente()
    {
        lblLog.Text = ("Atualizando tabela de clientes na base de dados...");

        DateTime dia = DateTime.Now; // DATA + HORA ATUAL

        var hora = dia.ToString("HH':'mm':'ss");
        try
        {
            ClienteService clienteService = new ClienteService();

            List<Cliente> listClientes = clienteService.GetClientesAtualizar();

            foreach (var clt in listClientes)
            {
                string address = clt.Endereco;
                
                string requestUri = string.Format("https://maps.googleapis.com/maps/api/geocode/xml?key={1}&address={0}&sensor=false", Uri.EscapeDataString(address), "chaveApi");

                WebRequest request = WebRequest.Create(requestUri);
                WebResponse response = request.GetResponse();
                XDocument xdoc = XDocument.Load(response.GetResponseStream());

                XElement result = xdoc.Element("GeocodeResponse").Element("result");
                if (result != null)
                {
                    XElement locationElement = result.Element("geometry").Element("location");
                    XElement lat = locationElement.Element("lat");
                    XElement lng = locationElement.Element("lng");

                    Double resultLat = Convert.ToDouble(lat.Value, System.Globalization.CultureInfo.InvariantCulture);
                    Double resultLng = Convert.ToDouble(lng.Value, System.Globalization.CultureInfo.InvariantCulture);

                    clt.Latitude = resultLat;
                    clt.Longitude = resultLng;
                    clienteService.AtualizaPosicaoCliente(clt.Cod_cliente, clt.Latitude, clt.Longitude);
                }
                else
                {
                    clienteService.AtualizaPosicaoCliente(clt.Cod_cliente, 0, 0);
                }
            }
        }
        catch (MySql.Data.MySqlClient.MySqlException ex)
        {
        }
        lblLog.Text = ("Atualização de Clientes finalizada as " + hora);
        timerAttPos.Enabled = true;


    }


    private void SalvaLogErro()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage excel = new ExcelPackage())
        {
            excel.Workbook.Worksheets.Add("Plan1");
            excel.Workbook.Worksheets.Add("Plan2");
            excel.Workbook.Worksheets.Add("Plan3");

            var headerRow = new List<string[]>()
                {
                new string[] { "Função", "Mensagem de erro", "Data e hora do erro"}
                };

            // Determina o range dos títulos em colunas (e.g. A1:D1)
            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

            // Seleciona uma planilha
            var worksheet = excel.Workbook.Worksheets["Plan1"];

            // Insere os dados dos títulos no range definido acima.
            worksheet.Cells[headerRange].LoadFromArrays(headerRow);

            List<Erro> cellData = new List<Erro>();

            worksheet.Cells[2, 1].LoadFromCollection(cellData);

            Directory.CreateDirectory(@"C:\Log");

            FileInfo excelFile = new FileInfo(@"C:\Log\teste.xlsx");
            excel.SaveAs(excelFile);
        }
    }

    public static string RemoveDiacritics(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] > 255)
                sb.Append(text[i]);
            else
                sb.Append(s_Diacritics[text[i]]);
        }

        return sb.ToString();
    }

    private static readonly char[] s_Diacritics = GetDiacritics();
    private static char[] GetDiacritics()
    {
        char[] accents = new char[256];

        for (int i = 0; i < 256; i++)
            accents[i] = (char)i;

        accents[(byte)'á'] = accents[(byte)'à'] = accents[(byte)'ã'] = accents[(byte)'â'] = accents[(byte)'ä'] = 'a';
        accents[(byte)'Á'] = accents[(byte)'À'] = accents[(byte)'Ã'] = accents[(byte)'Â'] = accents[(byte)'Ä'] = 'A';

        accents[(byte)'é'] = accents[(byte)'è'] = accents[(byte)'ê'] = accents[(byte)'ë'] = 'e';
        accents[(byte)'É'] = accents[(byte)'È'] = accents[(byte)'Ê'] = accents[(byte)'Ë'] = 'E';

        accents[(byte)'í'] = accents[(byte)'ì'] = accents[(byte)'î'] = accents[(byte)'ï'] = 'i';
        accents[(byte)'Í'] = accents[(byte)'Ì'] = accents[(byte)'Î'] = accents[(byte)'Ï'] = 'I';

        accents[(byte)'ó'] = accents[(byte)'ò'] = accents[(byte)'ô'] = accents[(byte)'õ'] = accents[(byte)'ö'] = 'o';
        accents[(byte)'Ó'] = accents[(byte)'Ò'] = accents[(byte)'Ô'] = accents[(byte)'Õ'] = accents[(byte)'Ö'] = 'O';

        accents[(byte)'ú'] = accents[(byte)'ù'] = accents[(byte)'û'] = accents[(byte)'ü'] = 'u';
        accents[(byte)'Ú'] = accents[(byte)'Ù'] = accents[(byte)'Û'] = accents[(byte)'Ü'] = 'U';

        accents[(byte)'ç'] = 'c';
        accents[(byte)'Ç'] = 'C';

        accents[(byte)'ñ'] = 'n';
        accents[(byte)'Ñ'] = 'N';

        accents[(byte)'ÿ'] = accents[(byte)'ý'] = 'y';
        accents[(byte)'Ý'] = 'Y';

        return accents;
    }

    private void timerAttVeic_Tick(object sender, EventArgs e)
    {
        segundoVeic--;

        if (segundoVeic < 0)
        {
            segundoVeic = 59;
            minutoVeic--;
            if (minutoVeic < 0)
            {
                minutoVeic = 59;
                horaVeic--;
            }
        }

        //lblTimerVeic.Text = horaVeic + ":" + minutoVeic + ":" + segundoVeic;
        if (horaVeic == 0 && minutoVeic == 0 && segundoVeic == 0)
        {
            timerAttVeic.Enabled = false;
            AtualizaVeic();
            horaVeic = 24;
            minutoVeic = 0;
            segundoVeic = 0;
        }
    }

    private void timerAttPos_Tick(object sender, EventArgs e)
    {

        segundoPos--;
        if (minutoPos > 0)
        {
            if (segundoPos < 0)
            {
                segundoPos = 59;
                minutoPos--;
            }
        }

        lblTimerPos.Text = horaPos + ":" + minutoPos + ":" + segundoPos;
        if (horaPos == 0 && minutoPos == 0 && segundoPos == 0)
        {
            timerAttPos.Enabled = false;
            AtualizaPos();
            minutoPos = 15;
            segundoPos = 0;

        }
    }

    private void Form1_Load(object sender, EventArgs e)
    {

    }
}
}
