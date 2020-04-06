# sharepoint-event-receiver

O código possui um projeto SharePoint 2013 Server (On-Premisses) com recurso de Event Receiver. Foi utilizado para resolução
de tratamentos de eventos em uma lista chamada Solicitações existente em um site Sharepoint. Tal solução atendeu um cenário real em 
uma empresa de geração de energia.

Ao adicionar uma nova solicitação ou altera-lá, é enviado emails ao solicitante e  aos facilitadores que são usuários vindos de campos
do tipo People Picker. E Event Receiver tratam momentos em que as solicitações são atualizadas bloqueando determinadas alteração 
seguindo seus status. Segue também a regra de bloquear a exclusão da solicitação conforme determinado status.

No código você encontra:

- Projeto SharePoint 2013 Server criado no Visual Studio (Community)
- Criação de Featue a nível Site para ativação/desativação do Event Receiver
- Linguagem C# aplicada a plataforma .NET Framework 4.5
- MVC básico com a pasta Model organizando os ojetos e Controller realizando chamada de funções
- Envio de mail utilzando o SMTP SharePoint de forma programatica


