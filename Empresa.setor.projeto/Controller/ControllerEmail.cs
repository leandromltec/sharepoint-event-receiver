using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Empresa.setor.projeto.Controller;
using Empresa.setor.projeto;
using Microsoft.SharePoint;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

namespace Empresa.setor.projeto.Controller
{
    class ControllerEmail
    {
        /// <summary>
        /// Mensagem de email a ser enviada
        /// </summary>
        /// <param name="usuarioSolicitante"></param>
        /// <param name="assunto"></param>
        /// <param name="solicitacao"></param>
        public static void mensagemEmailNovaSolicitacao(SPWeb web, SPUser usuario, SPUser solicitante, string status, string assunto, ref SPListItem solicitacao)
        {

            StringBuilder mensagem = new StringBuilder();

            mensagem.Append("<div style='font-size:15px'><b>Prezado (a)</b>");
            mensagem.Append("</br>");

            if (status == "EM ANÁLISE")
            {
                mensagem.Append("</br>");
                mensagem.Append("<b>Uma nova Solicitação de Serviço foi criada</b>");
            }

            if (status == "DEMANDA INCOMPATÍVEL")
            {

                mensagem.Append("</br>");
                mensagem.Append("<b>Uma Solicitação de Serviço é Incompatível</b>");

            }

            if (status == "EM ATENDIMENTO")
            {
                mensagem.Append("</br>");
                mensagem.Append("<b>Uma Solicitação de Serviço está em Atendimento</b>");

            }

            if (status == "ATENDIMENTO CONCLUÍDO")
            {

                mensagem.Append("</br>");
                mensagem.Append("<b>UMA SOLICITAÇÃO FOI CONCLUÍDA</b>");

            }

            if (status == "ATENDIMENTO CANCELADO")
            {

                mensagem.Append("</br>");
                mensagem.Append("<b>Uma Solicitação de Serviço foi cancelada</b>");
            }


            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<b>Código:</b> " + solicitacao["Title"].ToString());
            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<b>Solicitante: </b> " + solicitante.Name);
            mensagem.Append("</br>");

            if (solicitacao["Matricula"].ToString() != "")
            {
                mensagem.Append("<b>Matrícula: </b>" + solicitacao["Matricula"].ToString());
            }


            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<b>Status:</b> " + solicitacao["Status"].ToString());
            mensagem.Append("</br>");
            mensagem.Append("</br>");

            //Converte para DateTime para obter apenas a data do campo Prazo de Atendimento
            DateTime dataPRazoAtendimento = DateTime.Parse(solicitacao["PrazoFimAtendimento"].ToString());
            mensagem.Append("<b>Prazo para Atendimento:</b> " + dataPRazoAtendimento.Day.ToString("00") + "/" + dataPRazoAtendimento.Month.ToString("00") + "/" + dataPRazoAtendimento.Year.ToString("00"));

            if (status == "DEMANDA INCOMPATÍVEL" || status == "ATENDIMENTO CANCELADO")
            {
                string fraseJustificativa = string.Empty;
                string fraseUsuario = string.Empty;

                if (status == "DEMANDA INCOMPATÍVEL")
                {
                    fraseJustificativa = "Justificativa da Incompatibilidade: ";
                    fraseUsuario = "Usuário que definiu a Incompatibilidade: ";
                }

                if (status == "ATENDIMENTO CANCELADO")
                {
                    fraseJustificativa = "Justificativa do Cancelamento: ";
                    fraseUsuario = "Usuário que Cancelou a solicitação: ";
                }

                mensagem.Append("</br>");
                mensagem.Append("</br>");
                mensagem.Append("<b>" + fraseJustificativa + "</b>");
                mensagem.Append("</br>" + solicitacao["AtJustificativa"].ToString());
                mensagem.Append("</br>");
                mensagem.Append("</br>");
                mensagem.Append("<b>" + fraseUsuario + "</b> " + web.CurrentUser.Name);
                mensagem.Append("</br>");
                string[] matriculaUsuarioCorrente = web.CurrentUser.LoginName.Split('\\');
                mensagem.Append("<b>Matrícula: </b> " + matriculaUsuarioCorrente[1].ToUpper());

            }

            if (status == "EM ATENDIMENTO" || status == "ATENDIMENTO CONCLUÍDO")
            {
                mensagem.Append("</br>");
                mensagem.Append("</br>");

                SPFieldUserValue responsavelServico = new SPFieldUserValue(web, solicitacao["ResponsavelPeloServi_x00e7_o"].ToString());

                mensagem.Append("<b>Responsável pelo Serviço: </b>" + responsavelServico.User.Name);
                mensagem.Append("</br>");

                string[] matriculaResponsavelServico = responsavelServico.User.LoginName.Split('\\');
                mensagem.Append("<b>Matrícula: </b> " + matriculaResponsavelServico[1].ToUpper());

            }

            if (status == "ATENDIMENTO CONCLUÍDO")
            {
                mensagem.Append("</br>");
                mensagem.Append("</br>");
                string urlImagem = web.Url + "/PublishingImages/icone_pesquisa_qualidade.png";

                mensagem.Append("<a href=" + web.Url + "/SitePages/Avaliacao.aspx?IDSolicitacao=" + solicitacao.ID + "?CodigoSolicitacao=" + solicitacao["Title"].ToString().Trim() + " style='cursor:pointer'><img src='" + urlImagem + "' alt='Pesquisa' /></a>");

            }

            mensagem.Append("</br>");
            mensagem.Append("</br>");
            mensagem.Append("<a href=" + web.Url + "/Lists/SolicitacaoDeServico/DispForm.aspx?ID=" + solicitacao.ID + ">Clique no link para acessar a solicitação</a></div>");
            mensagem.Append("</br>");
            mensagem.Append("</br>");


            //Verifica se existe usuário e se o mesmo possui email 
            if (usuario.Email != null)
            {
                if (usuario.Email.ToString() != "")
                {
                    envioEmail(web, usuario.Email.ToLower().Trim(), assunto, mensagem.ToString());
                }
            }

        }


        /// <summary>
        /// Email a ser enviando para os usuários do campo Facilitadores (item Área da lista Áreas)
        /// </summary>
        /// <param name="item"></param>
        /// <param name="email"></param>
        /// <param name="assunto"></param>
        /// <param name="mensagem"></param>
        /// <param name="link"></param>
        public static void envioEmail(SPWeb web, string email, string assunto, string mensagem)
        {
            Email envioEmail = new Email();

            envioEmail.Para = email;

            envioEmail.Assunto = assunto;
            envioEmail.Corpo = mensagem.ToString();

            envioEmail.EnviaEmail(web);
        }

        /// <summary>
        /// Envio o email para o Solicitante (usuário do campo Solicitante do Solicitação criada)
        /// </summary>
        public static void emailSolicitante(SPWeb web, SPListItem solicitacao, string status, string assuntoEmail)
        {
            //Informações do Usuário no campo Solicitante
            SPFieldUserValue solicitanteEmail = new SPFieldUserValue(web, solicitacao["Solicitante"].ToString());

            //Envia email para o solicitante
            ControllerEmail.mensagemEmailNovaSolicitacao(web, solicitanteEmail.User, solicitanteEmail.User, status, assuntoEmail, ref solicitacao);

        }



        /// <summary>
        /// Usuários que irão receber o determinado email (campo Facilitadores do item Área)
        /// </summary>
        /// <param name="web"></param>
        /// <param name="itemArea"></param>
        public static void enviarEmailUsuariosFacilitadores(SPWeb web, string status, SPListItem itemArea, SPListItem solicitacao, string assuntoEmail)
        {

            SPFieldUserValueCollection colecaoUsuariosFacilitadores = ControllerFacilitadores.usuariosFacilitadores(web, itemArea["Facilitadores"].ToString());


            //Busca informações de usuários no coleção de informações de facilitadores
            foreach (SPFieldUserValue usuarioFacilitador in colecaoUsuariosFacilitadores)
            {
                if (usuarioFacilitador.User.Email != "" || usuarioFacilitador.User != null)
                {

                    SPFieldUserValue solicitanteEmail = new SPFieldUserValue(web, solicitacao["Solicitante"].ToString());



                    ControllerEmail.mensagemEmailNovaSolicitacao(web, usuarioFacilitador.User, solicitanteEmail.User, status, assuntoEmail, ref solicitacao);

                }
            }
        }


        public static string emailDeEnvio(SPWeb web)
        {
            SPList listaEmailSolicitacaoServico = web.Lists["E-mail para Envio de Solicitações de Serviço"];
            SPListItemCollection colecaoEmailSolicitacaoServico = listaEmailSolicitacaoServico.GetItems();

            string emailDe = null;

            if (colecaoEmailSolicitacaoServico.Count > 0)
            {
                emailDe = colecaoEmailSolicitacaoServico[0]["Title"].ToString().ToLower().Trim();


            }
            return emailDe;
        }


    }
}
