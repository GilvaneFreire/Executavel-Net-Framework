using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class Cliente
    {
        public string Cod_cliente { get; set; }
        public string Nome { get; set; }
        public string Endereco { get; set; }
        public double Latitude { get; set; }
        public double Longitude { get; set; }
        public string DataAtualizacaoCliente { get; set; }

        public static explicit operator Cliente(bool v)
        {
            throw new NotImplementedException();
        }
    }
}
