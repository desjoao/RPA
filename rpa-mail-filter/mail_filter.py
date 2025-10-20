import os
import re
import pandas
import base64
import traceback
from utils import Utils
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class MailFilter:
    """Classe que agrupa métodos e atributos referentes à automação de e-mails de candidaturas
    de vagas de trabalho.
    TODO: Generalizar a classe para automatizar quaisquer tipos de e-mail dentro de um determinado assunto.
    """
    def __init__(self):
        self.__version__ = "0.1.0.0"
        self.nome_rpa= 'rpa-mail-filter'
        self.utils = Utils(self.nome_rpa)
        self.config = self.utils.carrega_json('config.mail')
        self.planilha = self.config['planilha']['caminho']
        self.path_pastas = self.config['pastas_candidatos']['caminho']
        self.service = ''
    
    def main(self):
        """Fluxo da automação, opera de forma a não jogar uma exceção na tela do usuário em casos de falha.
        """
        try:
            if not self.carrega_dados_Google():
                self.utils.log_fim()
                exit()
            if not self.autenticar_gmail():
                self.utils.log_fim()
                exit()
            emails = self.buscar_gmails()
            if not emails:
                self.utils.log_fim()
                exit()
            if not self.extrair_gmails(emails):
                self.utils.log_fim()
                exit()
            msg = 'E-mails processados com sucesso.'
            print(msg)
            self.utils.log_acerto(msg)
            self.utils.log_fim()
            exit()
        
        except Exception as e:
            msg = f'Erro não mapeado: {str(e)}.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            self.utils.log_fim()
            exit()
        
    def carrega_dados_Google(self):
        """Carrega os dados referentes à API da Google, contidos no arquivo config.mail.json
        """
        try:
            self.gtoken = self.config["API"]["Google"]["token_file"]
            self.gcredentials = self.config["API"]["Google"]["credentials_file"]
            self.gscopes = self.config["API"]["Google"]["scopes"]
            return True
        except Exception as e:
            msg = f"Erro ao carregar dados da API do gmail. \nTraceback: {traceback.format_exc()}"
            print(msg)
            self.utils.log_erro(msg)
            return False
    
    def autenticar_gmail(self):
        """Busca fazer a conexão com a conta do usuário no Google.
        Atualmente o usuário deve ser adicionado como testador da aplicação
        para conseguir se conectar com a API.
        """
        try:
            creds = None
            if os.path.exists(self.gtoken):
                creds = Credentials.from_authorized_user_file(self.gtoken, self.gscopes)

            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.gcredentials, self.gscopes)
                    creds = flow.run_local_server(port=0)
                
                with open(self.gtoken, 'w') as token:
                    token.write(creds.to_json())
            self.service = build('gmail', 'v1', credentials=creds)
            return True
        
        except Exception as e:
            msg = f'Erro ao autenticar conta no gmail.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return False
    
    def buscar_gmails(self):
        """Busca e-mails não lidos que contenham o assunto passado por meio do config.mail.json
        """
        try:
            assunto = self.config["ajuste_fino"]["assunto_email"]
            lote = int(self.config["ajuste_fino"]["tamanho_lote"])
            consulta = f'subject:{assunto} is:unread'
            
            resultado = self.service.users().messages().list(userId='me', q=consulta, maxResults=lote).execute()
            if resultado['resultSizeEstimate'] == 0:
                print('Não há nenhum e-mail de candidatura de emprego.')
                self.utils.log_fim()
                exit()
            mensagens = resultado.get('messages', [])
            return mensagens
        
        except Exception as e:
            msg = f'Erro ao buscar e-mails pelo assunto informado: {assunto}.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return None
    
    def extrair_gmails(self, mensagens):
        """Faz a extração do corpo da mensagem do e-mail e dos arquivos em anexo (se houverem).
        Processa os dados extraídos.
        """
        try:
            for msg in mensagens:
                msg_id = msg['id']
                mensagem = self.service.users().messages().get(userId='me', id=msg_id).execute()
                payload = mensagem['payload']

                corpo = self.extrair_corpo(payload)
                anexos = self.extrair_anexos(payload)
                
                # Não interrompe o fluxo caso o processamento em algum email não obtenha sucesso.
                if not self.processa_dados(msg_id, corpo, anexos):
                    msg = f'Falha ao processar e-mail {msg_id}.'
                    self.utils.log_erro(msg)
                    continue 
                
                self.marcar_email(msg_id)
                print(f'Email {msg_id} processado com sucesso.')
            return True
        
        except Exception as e:
            msg = f'Erro ao extrair emails.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return False
    
    def extrair_corpo(self, payload):
        try:
            if payload.get('body', {}).get('data'):
                return self.decodificar_dado(payload['body']['data'])
            if 'parts' in payload:
                for parte in payload['parts']:
                    mime = parte.get('mimeType', '')
                    body = parte.get('body', {})
                    data = body.get('data')

                    if mime == 'text/plain' and data:
                        return self.decodificar_dado(data)
                    elif mime == 'text/html' and data:
                        html = self.decodificar_dado(data)
                        return re.sub('<[^<]+?', '', html)
                    if 'parts' in parte:
                        resultado = self.extrair_corpo(parte)
                        if resultado != "":
                            return resultado
                return '(vazio)'

        except Exception as e:
            msg = f'Erro ao extrair corpo do e-mail\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return None
    
    def extrair_anexos(self, payload):
        try:
            anexos = []
            if 'parts' in payload:
                anexos = self.buscar_anexos(payload['parts'], anexos)
            return anexos

        except Exception as e:
            msg = f'Erro ao extrair arquivos anexados ao email.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return None
    
    def buscar_anexos(self, partes, anexos):
        for parte in partes:
            if 'filename' in parte and parte['filename']:
                if 'attachmentId' in parte.get('body', {}):
                    anexos.append({
                        'filename': parte['filename'],
                        'attachmentId': parte['body']['attachmentId']
                    })
            if 'parts' in parte:
                anexos = self.buscar_anexos(parte['parts'], anexos)
        return anexos
    
    def decodificar_dado(self, dado):
        return base64.urlsafe_b64decode(dado.encode('UTF-8')).decode('utf-8', errors='ignore')
    
    def processa_dados(self, msg_id, corpo, anexos):
        """Retira os valores referentes ao Nome, Telefone e Vaga informados pelo candidato no e-mail."""
        try:
            data = list(corpo.split('\n'))
            for i in range(0, len(data)):
                if 'Nome' in data[i]:
                    nome = list(data[i].split(':'))
                    nome = nome[1][1:-1]
                elif 'Telefone' in data[i]:
                    telefone = list(data[i].split(':'))
                    telefone = telefone[1][1:-1]
                elif 'Vaga' in data[i]:
                    vaga = list(data[i].split(':'))
                    vaga = vaga[1][1:-1]
                else:
                    continue
            
            # Se o e-mail não tiver anexo, loga o erro e passa para o próximo e-mail
            if len(anexos) == 0:
                msg = f'E-mail do candidato(a) {nome} foi desconsiderado, pois não possui anexo.'
                self.utils.log_erro(msg)
                return True
            
            self.atualizar_planilha(nome, telefone, vaga)
            self.baixar_anexo(msg_id, nome, anexos)
            return True

        except Exception as e:
            return False
    
    def baixar_anexo(self, msg_id, nome, anexos):
        try:
            if not os.path.exists(self.path_pastas):
                os.makedirs(self.path_pastas)
            
            for anexo in anexos:
                nome_anexo = anexo['filename']
                id_anexo = anexo['attachmentId']

                dado_anexo = self.service.users().messages().attachments().get(
                    userId='me',
                    messageId=msg_id,
                    id=id_anexo
                ).execute()
                dado = dado_anexo['data']
                file_data = base64.urlsafe_b64decode(dado.encode('UTF-8'))
            
                pasta_candidato = os.path.join(self.path_pastas, nome.upper())
                if not os.path.exists(pasta_candidato):
                    os.makedirs(pasta_candidato)
            
                caminho = os.path.join(pasta_candidato, nome_anexo)
                with open(caminho, 'wb') as f:
                    f.write(file_data)    
                    print(f'Anexo \'{nome_anexo}\' salvo com sucesso.')
            return True
        
        except Exception as e:
            msg = f'Erro ao salvar anexos do email: {msg_id}.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return False
    
    def atualizar_planilha(self, nome, telefone, vaga):
        try:
            if not os.path.exists(self.planilha):
                df = pandas.DataFrame(columns=['Nome', 'Telefone', 'Vaga'])
            else:
                df = pandas.read_excel(self.planilha)
            
            nova_linha = pandas.DataFrame([{
                'Nome': nome,
                'Telefone': telefone,
                'Vaga': vaga
            }])
            df = pandas.concat([df, nova_linha], ignore_index=True)
            print(df.head())
            df.to_excel(self.planilha, index=False)
            return True
            
        except Exception as e:
            msg = f'Erro ao atualizar planilha.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return False

    def marcar_email(self, msg_id):
        """Marca o e-mail como lido para que ele não volte a ser processado."""
        try:
            self.service.users().messages().modify(
                userId='me',
                id=msg_id,
                body={
                    'removeLabelIds': ['UNREAD']
                }
            ).execute()
            return True
        
        except Exception as e:
            msg = f'Erro ao marcar e-mail como lido.\nTraceback: {traceback.format_exc()}'
            print(msg)
            self.utils.log_erro(msg)
            return False

if __name__ == "__main__":
    mail = MailFilter()
    mail.main()