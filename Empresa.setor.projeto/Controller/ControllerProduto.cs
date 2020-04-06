using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

/* Classe responsável pelo Produto */

namespace Empresa.setor.projeto.Controller
{
    public class ControllerProduto
    {
        //Busca o produto
        public static SPListItem BuscaProduto(SPWeb web, string produto)
        {
            //Lista de Produtos
            SPList produtolist = web.Lists["Produtos"];

            //Instancia de uma nova Query
            SPQuery queryContentyProduto = new SPQuery();

            //Busca na lista de Produtos pelo nome do tipod e conteudo (Content Type) do item que foi adicionado no Acompanhamento das Solicitações de Serviço
            queryContentyProduto.Query = "<Where><Eq><FieldRef Name='tipoconteudo'  /><Value Type='Text' >" + produto.Trim() + "</Value></Eq></Where>";

            //Retorna apenas a um linha (um único item - um resultado)
            queryContentyProduto.RowLimit = 1;

            //Busca uma coleçaõ de produtos utilizando a SPQuery
            SPListItemCollection colecaoProdutos = produtolist.GetItems(queryContentyProduto);

            //Retorna ao primeiro item do indice da colecao de Produtos encontrados
            SPListItem itemProduto = colecaoProdutos[0];

            //Retorna ao item Produto encontrado
            return itemProduto;

        }

    }
}
