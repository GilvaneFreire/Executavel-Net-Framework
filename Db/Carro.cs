using System;

namespace WindowsFormsApp1
{
    public class Carro
    {
        public int IdVeiculo { get; set; }
        public string Placa { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public DateTime DataPacote { get; set; }

        public string CidadeAtual { get; set; }
        public string UFatual { get; set; }

    }
}