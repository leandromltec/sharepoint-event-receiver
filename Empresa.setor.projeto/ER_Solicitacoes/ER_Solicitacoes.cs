using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Empresa.setor.projeto.Model;
using Empresa.setor.projeto.Controller;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Linq;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

/* Eventos são tratados no momento que uma solicitação é adicionada, editada ou na tentativa de sua exclusão */

namespace Empresa.setor.projeto.ER_Solicitacoes
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class ER_Solicitacoes : SPItemEventReceiver
    {
        /// <summary>
        /// Momento que a solicitação está sendo atualizada
        /// </summary>
        public override void ItemUpdating(SPItemEventProperties properties)
        {
            base.ItemUpdating(properties);

            this.EventFiringEnabled = false;

            if (properties.List.RootFolder.ServerRelativeUrl == "/Lists/Solicitacoes")
            {
                string status = properties.ListItem["Status"].ToString().ToUpper().Trim();


                if (status == "DEMANDA INCOMPATÍVEL")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = " Não é permitido alterar a solicitação porque a mesma é uma demanda incompatível";

                }


                if (status == "ATENDIMENTO CANCELADO")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = " Não é permitido alterar a solicitação por que a mesma foi cancelada";

                }

                if (status == "ATENDIMENTO CONCLUÍDO")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = " Não é permitido alterar a solicitação por que a mesma já foi concluída";

                }

                if (properties.AfterProperties["Status"].ToString().ToUpper() == "DEMANDA INCOMPATÍVEL")
                {
                    //Busca o item area para encontra os usuários Facilitadores que receberão o email pra nova solicitação
                    SPListItem itemProduto = ControllerProduto.BuscaProduto(properties.Web, properties.ListItem.ContentType.Name);

                    SPListItem itemArea = ControllerArea.BuscaArea(properties.Web, itemProduto);

                    SPFieldUserValueCollection colecaoUsuariosFacilitadores = ControllerFacilitadores.usuariosFacilitadores(properties.Web, itemArea["Facilitadores"].ToString());

                    bool contemFacilitador = false;

                    foreach (SPFieldUserValue usuarioFacilitador in colecaoUsuariosFacilitadores)
                    {

                        if (usuarioFacilitador.User.ID == properties.CurrentUserId)
                        {
                            contemFacilitador = true;

                        }

                    }

                    if (contemFacilitador == false)
                    {
                        properties.Status = SPEventReceiverStatus.CancelWithError;
                        properties.ErrorMessage = " Você não tem permissão para alterar o status para demanda incompatível. Apenas os falicitadores do produto podem realizar esta alteração";

                    }

                }


                if (status == "EM ANÁLISE")
                {

                    if (properties.AfterProperties["Status"].ToString().ToUpper() == "ATENDIMENTO CONCLUÍDO")
                    {

                        properties.Status = SPEventReceiverStatus.CancelWithError;
                        properties.ErrorMessage = " Não é permitido alterar a solicitação. Status da solicitação em análise poderá apenas ser alterada para status em atendimento";

                    }

                }


                if (properties.AfterProperties["Status"].ToString().ToUpper() == "EM ANÁLISE" && status != "EM ANÁLISE")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = " Não é permitido alterar a solicitação para status em análise porque a mesma se encontra  no status " + status.ToLower();

                }


                if (properties.AfterProperties["Status"].ToString().ToUpper() == "ATENDIMENTO CONCLUÍDO" && status != "EM ATENDIMENTO")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = " Não é permitido alterar a solicitação. Apenas solicitações em atendimento podem ser concluídas.";

                }


            }

            this.EventFiringEnabled = true;
        }


        /// <summary>
        /// Momento em que a solicitação está sendo deletada
        /// 
        /// </summary>
        public override void ItemDeleting(SPItemEventProperties properties)
        {
            base.ItemDeleting(properties);

            if (properties.List.RootFolder.ServerRelativeUrl == "/Lists/Solicitacoes")
            {
                string status = properties.ListItem["Status"].ToString().ToUpper().Trim();

                if (status == "DEMANDA INCOMPATÍVEL")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Não é permitido excluir esta solicitação por ser uma demanda incompatível.";
                }

                if (status == "ATENDIMENTO CANCELADO")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Não é permitido excluir esta solicitação por ter sido cancelada";
                }


                if (status == "ATENDIMENTO CONCLUÍDO")
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Não é permitido excluir esta solicitação porque a mesma já foi concluída.";
                }

            }

        }

        /// <summary>
        /// Após a solicitação ser adicionada
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            base.ItemAdded(properties);

            if (properties.List.RootFolder.ServerRelativeUrl == "/Lists/Solicitacoes")
            {
                try
                {
                    this.EventFiringEnabled = false;

                    //Gravar contentType Name no campo Produto 
                    properties.ListItem["Produto"] = properties.ListItem.ContentType.Name; properties.ListItem.Update();

                    //Criação do número de item 
                    SPListItem ItemAno = null;

                    SPList Controle = properties.Web.Lists["Controle de Código"];



                    //Retorna ao ultimo item da lista Controle de Código 
                    ItemAno = Controle.Items.OfType<SPListItem>().LastOrDefault();

                    if (ItemAno != null)
                    {
                        double proximo = 0;

                        //Função concatena o ano corrente com o último código somado a 1 (Ex: Último código 33 será 2020 (ano) + 34 = 2020.34) 
                        ControllerValidarAno.ValidaAno(ItemAno, ref Controle, ref proximo);

                        if (proximo != 1)
                        {
                            proximo = ((double)ItemAno["Último Código"]) + 1;
                            ItemAno["Último Código"] = proximo; ItemAno.Update();
                        }

                        //Atualiza o campo código concateando o ano atual após quatro posições adionando 0 
                        properties.ListItem["Código"] = DateTime.Now.Year.ToString() + "." + proximo.ToString().PadLeft(4, '0');
                        properties.ListItem.Update();

                        //Busca o item area para encontra os usuários Facilitadores que receberão o email pra nova solicitação 
                        SPListItem itemProduto = ControllerProduto.BuscaProduto(properties.Web, properties.ListItem.ContentType.Name);

                        //Busca área do respectivo produto 
                        SPListItem itemArea = ControllerArea.BuscaArea(properties.Web, itemProduto);

                        //Variável Global recebe o título concatenado com código para envio de email tanto para Solicitante quanto para 
                        string assuntoEmail = "Nova Solicitação de Serviço Criada: Código - " + properties.ListItem["Código"].ToString();
                        string status = properties.ListItem["Status"].ToString().ToUpper().Trim();

                        //Envia email para o Solicitante 
                        ControllerEmail.emailSolicitante(properties.Web, properties.ListItem, status, assuntoEmail);

                        //Envia email para os usuários do campo Facilitadores (item Área na lista Áreas) 
                        ControllerEmail.enviarEmailUsuariosFacilitadores(properties.Web, status, itemArea, properties.ListItem, assuntoEmail);

                        this.EventFiringEnabled = true;
                    }
                }


                catch (Exception e)
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Erro ao cadastrar solicitação." + e.Message;
                }

            }
        }

        /// <summary>
        /// Após a solicitação ser atualizada
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            base.ItemUpdated(properties);

            if (properties.List.RootFolder.ServerRelativeUrl == "/Lists/Solicitacoes")
            {

                try
                {
                    this.EventFiringEnabled = false;

                    if (properties.ListItem["Status"].ToString().ToUpper().Trim() != "EM ANÁLISE")
                    {

                        InformacoesEmail informacoesEmail = new InformacoesEmail();

                        //Busca o item area para encontra os usuários Facilitadores que receberão o email pra nova solicitação
                        SPListItem itemProduto = ControllerProduto.BuscaProduto(properties.Web, properties.ListItem.ContentType.Name);

                        SPListItem itemArea = ControllerArea.BuscaArea(properties.Web, itemProduto);


                        string status = properties.ListItem["Status"].ToString().ToUpper().Trim();

                        if (status == "ATENDIMENTO CONCLUÍDO")
                        {
                            informacoesEmail.assuntoEmail = "Solicitação de Serviço Concluída: Código - " + properties.ListItem["Código"].ToString();


                            //Envia email para o Solicitante
                            ControllerEmail.emailSolicitante(properties.Web, properties.ListItem, status, informacoesEmail.assuntoEmail);

                        }

                        if (status == "DEMANDA INCOMPATÍVEL")
                        {
                            informacoesEmail.assuntoEmail = "Solicitação de Serviço Incompatível: Código - " + properties.ListItem["Código"].ToString();


                            //Envia email para o Solicitante
                            ControllerEmail.emailSolicitante(properties.Web, properties.ListItem, status, informacoesEmail.assuntoEmail);
                        }


                        if (status == "ATENDIMENTO CANCELADO")
                        {
                            informacoesEmail.assuntoEmail = "Solicitação de Serviço Cancelada: Código - " + properties.ListItem["Código"].ToString();


                            //Envia email para o Solicitante
                            ControllerEmail.emailSolicitante(properties.Web, properties.ListItem, status, informacoesEmail.assuntoEmail);

                            //Envia email para os usuários do campo Facilitadores (item Área na lista Áreas)
                            ControllerEmail.enviarEmailUsuariosFacilitadores(properties.Web, status, itemArea, properties.ListItem, informacoesEmail.assuntoEmail);
                        }


                        if (status == "EM ATENDIMENTO")
                        {
                            informacoesEmail.assuntoEmail = "Solicitação de Serviço em Atendimento: Código - " + properties.ListItem["Código"].ToString();


                            //Envia email para o Solicitante
                            ControllerEmail.emailSolicitante(properties.Web, properties.ListItem, status, informacoesEmail.assuntoEmail);

                            //Envia email para os usuários do campo Facilitadores (item Área na lista Áreas)
                            ControllerEmail.enviarEmailUsuariosFacilitadores(properties.Web, status, itemArea, properties.ListItem, informacoesEmail.assuntoEmail);
                        }
                    }
                }

                catch (Exception e)
                {

                    properties.Status = SPEventReceiverStatus.CancelWithError;
                    properties.ErrorMessage = "Erro ao atualizar a solicitação." + e.Message;
                }

            }

        }

    }
}