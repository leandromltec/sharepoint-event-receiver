using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

/* Desenvolvido por Leandro M. Loureiro */
/* Linkedin - https://www.linkedin.com/in/leandro-loureiro-9921b927/ */

/* Classe responsável pelo construção e envio do Email */

namespace Empresa.setor.projeto
{

    public class Email
    {
        //Propriedades do Email
        public string Assunto { get; set; }
        public string Corpo { get; set; }
        public string Para { get; set; }
        public string CC { get; set; }
        public string CCO { get; set; }
        public string NomePortal { get; set; }

        private List<string> _Anexos;

        public List<string> Anexos
        {
            get
            {
                if (_Anexos == null)
                    _Anexos = new List<string>();
                return _Anexos;
            }
            set { _Anexos = value; }
        }


        public bool EnviaEmail(SPWeb Web)
        {
            //Armazena o endereço email do configurando no serviço de Email do Farm SharePoint
            string SharePointSmtp;
            string SharePointDe;

            try
            {
                SharePointSmtp = SPAdministrationWebApplication.Local.OutboundMailServiceInstance.Server.Address;
                SharePointDe = SPAdministrationWebApplication.Local.OutboundMailSenderAddress;

            }
            catch
            {
                return false;
            }

            //Instância para criação de um novo email
            MailMessage eMail = new MailMessage();

            //Propriedades do email 
            eMail.Subject = this.Assunto;
            eMail.Body = this.Corpo;
            eMail.IsBodyHtml = true;
            eMail.From = new MailAddress(SharePointDe, NomePortal);

            //Carrega os anexos
            foreach (string anexo in this.Anexos)
            {

                SPFile file = Web.GetFile(anexo);
                Attachment attach = new Attachment(file.OpenBinaryStream(), System.Net.Mime.MediaTypeNames.Application.Octet);
                attach.Name = file.Name;
                System.Net.Mime.ContentDisposition disposition = attach.ContentDisposition;
                disposition.CreationDate = file.TimeCreated;
                disposition.ModificationDate = file.TimeLastModified;
                disposition.ReadDate = file.TimeLastModified;
                eMail.Attachments.Add(attach);
            }

            //Verifique se existe destinário para o email
            if (!String.IsNullOrEmpty(this.Para))
            {
                string[] destinatariosPara = this.Para.Split(';');

                foreach (string destinatarioPara in destinatariosPara)
                {
                    if (!String.IsNullOrEmpty(destinatarioPara))
                    {
                        MailAddress DestinatarioPara = new MailAddress(destinatarioPara);
                        eMail.To.Add(DestinatarioPara);
                    }
                }
            }

            //Verifique se existe cópia para o email
            if (!String.IsNullOrEmpty(this.CC))
            {
                string[] destinatariosCC = this.CC.Split(';');

                foreach (string destinatarioCC in destinatariosCC)
                {
                    if (!String.IsNullOrEmpty(destinatarioCC))
                    {
                        MailAddress DestinatarioCC = new MailAddress(destinatarioCC);
                        eMail.CC.Add(DestinatarioCC);
                    }
                }
            }

            //Verifique se existe cópia oculta para o email
            if (!String.IsNullOrEmpty(this.CCO))
            {
                string[] destinatariosCCO = this.CCO.Split(';');

                foreach (string destinatarioCCO in destinatariosCCO)
                {
                    if (!String.IsNullOrEmpty(destinatarioCCO))
                    {
                        MailAddress DestinatarioCCO = new MailAddress(destinatarioCCO);
                        eMail.Bcc.Add(DestinatarioCCO);
                    }
                }
            }

            //Instância de smpt do SharePoint
            SmtpClient smtpClient = new SmtpClient(SharePointSmtp);

            //Envia email com as informações do objeto email
            smtpClient.Send(eMail);

            return true;
        }
    }
}

