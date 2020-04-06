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
    public class ControllerValidarAno
    {

        public static void ValidaAno(SPListItem ItemAno, ref SPList Controle, ref double Proximo)
        {
            if (Convert.ToString(ItemAno["Ano"]) != DateTime.Now.Year.ToString())
            {
                Proximo = 1;
                SPListItem itemControle = Controle.AddItem();
                itemControle["Ano"] = DateTime.Now.Year.ToString();
                itemControle["Último Código"] = 1;
                itemControle.Update();
            }
            else
                Proximo = 0;
        }
    }

    
}
