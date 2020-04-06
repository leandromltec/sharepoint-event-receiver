using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

namespace Empresa.setor.projeto.Controller
{
    public class ControllerArea
    {
        //Busca área relacionada a o produto      
        public static SPListItem BuscaArea(SPWeb web, SPListItem produto)
        {

            SPFieldLookupValue area = new SPFieldLookupValue(produto["Area"].ToString());

            SPList listaArea = web.Lists["Areas"];
            SPListItem itemArea = listaArea.GetItemById(area.LookupId);

            //Retorna ao item área encontrado
            return itemArea;
        }
    }
}
