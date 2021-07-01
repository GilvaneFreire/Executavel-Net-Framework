using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;

namespace WindowsFormsApp1
{
    public class ClienteService : Service
    {
        public List<Cliente> GetClientesAtualizar()
        {

            MySqlConnection conexao = GetConexao();
            List<Cliente> listClientes = new List<Cliente>();

            String sql = @"select distinct cli.cod_cliente,
			                                cli.nome as nome, 
		                                    CONCAT(rtrim(ltrim(cli.endereco)), ', ',
                                                           rtrim(ltrim(cli.numero)), ', ', 
                                                           rtrim(ltrim(cli.bairro)), ', ',
                                                           rtrim(ltrim(cli.cidade)), ', ', 
                                                           rtrim(ltrim(cli.Estado)), ', ',
                                                           rtrim(ltrim(cli.cep)), ', ', 
                                                             case when cli.pais is null = ''
                                                    then 'Brasil' else cli.pais end ) as endereco
                                                                    from  carga c
                                                                    join cliente cli on c.cod_destinatario = cli.cod_cliente
                                                                   where c.dt_carga >= '20190101'
                                                                     and cli.cidade <> 'EXTERIOR'
                                                                     and cli.cod_cliente not in (select cod_cliente 
                                                                                                  from cliente_aux 
                                                                                                 where lat is not null)";
            var command = new MySql.Data.MySqlClient.MySqlCommand(sql, conexao);
            MySqlDataReader rdr = command.ExecuteReader();

            while (rdr.Read())
            {
                Cliente cliente = new Cliente();
                var enderecoFormatado = (string)rdr["endereco"];
                cliente.Endereco = RemoveDiacritics(enderecoFormatado);
                cliente.Cod_cliente = (string)rdr["cod_cliente"];
                cliente.Nome = (string)rdr["nome"];
                listClientes.Add(cliente);
            }
            FecharConexao();
            return listClientes;

        }
        public bool AtualizaPosicaoCliente(String codCliente, Double? Lat, Double? Long)
        {
            String sql;
            if (GetCliente(codCliente) != null)
            {
                sql = @"UPDATE cliente_aux 
                           SET lat = @lat, 
                               lng = @lng, 
                               dataAtualizacaoCliente = @dataAtualizacaoCliente 
                         WHERE cod_cliente = @cod_cliente";

            }
            else
            {
                sql = @"INSERT INTO cliente_aux( 
                                        cod_cliente, 
                                        lat, 
                                        lng , 
                                        dataAtualizacaoCliente ) 
                                values (
                                        @cod_cliente, 
                                        @lat, 
                                        @lng , 
                                        @dataAtualizacaoCliente )";
            }
            MySqlConnection conexao = GetConexao();
            using (var command = new MySql.Data.MySqlClient.MySqlCommand(sql, conexao))
            {
                command.Prepare();
                command.Parameters.AddWithValue("@cod_cliente", codCliente);
                command.Parameters.AddWithValue("@lat", Lat);
                command.Parameters.AddWithValue("@lng", Long);
                command.Parameters.AddWithValue("@dataAtualizacaoCliente", DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss") );

                bool ret = command.ExecuteNonQuery() >= 0;
                FecharConexao();
                return ret;

            }
            
        }
        public Cliente GetCliente(String codCliente)
        {
            MySqlConnection conexao = GetConexao();
            String sql = @"SELECT * from cliente_aux where cod_cliente = @cod_cliente ";
            var command = new MySql.Data.MySqlClient.MySqlCommand(sql, conexao);
            command.Parameters.AddWithValue("@cod_cliente", codCliente);
            MySqlDataReader rdr = command.ExecuteReader();
            
            
            Cliente cliente = null;

            while (rdr.Read())
            {
                cliente = new Cliente();
                var enderecoFormatado = (string)rdr["endereco"];
                cliente.Endereco = RemoveDiacritics(enderecoFormatado);
                cliente.Cod_cliente = (string)rdr["cod_cliente"];
                cliente.Nome = (string)rdr["nome"];
                
            }
            FecharConexao();
            return cliente;

        }
        
                   
    }
}
