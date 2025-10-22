# RPA - Mail_Filter (Case Líber)

## 1. Sobre
A automação realiza uma filtragem e o processamento de e-mails recebidos para aplicações a vagas de emprego. Caso o e-mail venha em um formato prestabelecido, contendo nome, telefone, vaga de aplicação e documento de CV em anexo, a automação:

    1 -​ ​Lê os e-mails recebidos com o assunto: “Candidatura – [Nome da Vaga]”;​
    
    2 -​ ​Salva os arquivos PDF dos currículos em uma pasta específica com o​ nome do candidato;
    
    3 -​ ​Registra os dados do candidato (nome, vaga e telefone) em uma planilha​ do Excel.

## 2. Tecnologias utilizadas
A automação foi feita com Python com uso de API para o gmail. O motivo se deu pela vasta gama de bibliotecas disponíveis em Python para se criar um script de automação de processos, assim como pela liberdade de desenvolvimento que essa linguagem de programação proporciona em relação a ferramentas low-code/no-code. 

A automação pode ser utilizada sem a necessidade de um ambiente Python instalado na máquina do usuário, para isso sendo necessários apenas o arquivo executável do RPA, os arquivos de dados que são lidos na execução do RPA (json e xlsx) e uma conta cadastrada no gmail.

Por não possuir monitoramento ativo da caixa de entrada do e-mail do usuário, a automação deve ser programada para ser executada pelo agendador de tarefas em intervalos de tempo, definidos pelo usuário.

## 3. Desafios
A integração da automação com a API Restful dos clientes de e-mail (Google e Microsoft) se mostrou desafiadora em razão de ser uma novidade para mim. Com o tempo disponibilizado para o desafio, não tive êxito em implementar a automação no cliente Outlook.

## 4. Confiabilidade
Apesar de ser um protótipo, a automação opera bem dentro dos casos previstos, em que erros e falhas são tratados a não quebrarem a execução e causar perda de dados. Logs de execução são registrados a cada execução, como forma de se ter certo grau de auditoria sobre o RPA.

## 5. Melhorias
O RPA possui algumas melhorias mapeadas:

    1 - Retirar as inforamções 'Nome', 'Telefone' e 'Vaga' do corpo de e-mail que não estejam estruturados em um template.
    
    2 - Criar uma integração com a API do Outlook.

    3 - Criar um envio automático de resposta aos e-mails que não encaminharam o currículo em anexo.

    4 - Criar uma monitoração ativa da caixa de entrada, para que a automação seja executada sempre que um novo e-mail de candidatura de emprego chegue.

## 6. Informações
Para criar o arquivo executável dessa automação, é necessário utilizar o arquivo rpa_mail_filter.spec com o software PyInstaller.
