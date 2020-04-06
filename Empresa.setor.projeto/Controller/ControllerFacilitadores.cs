using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

/* Classe recebe coleção de informações de usuários do campo People Picker chamado de Facilitadores */

namespace Empresa.setor.projeto.Controller
{
    public class ControllerFacilitadores
    {
        public static SPFieldUserValueCollection usuariosFacilitadores(SPWeb web, string campoFacilitadores)
        {

            //Campos dos usuários Facilitadores
            SPFieldUserValueCollection colecaoUsuariosFacilitadores = new SPFieldUserValueCollection(web, campoFacilitadores);

            return colecaoUsuariosFacilitadores;

        }
    }
}
