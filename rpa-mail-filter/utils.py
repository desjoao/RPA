import os
import json
from datetime import datetime

class Utils:
    def __init__(self, nome_rpa: str):
        self.rpa = nome_rpa
        self.caminho_log = os.path.join(os.getcwd(), f'log_{self.rpa}.txt' )
        self.log_inicio()
    
    def carrega_json(self, nome:str) -> dict:
        try:
            config = json.load(open(nome+'.json', 'r'))
            return config
        except Exception as e:
            msg = f'Arquivo json não encontrado.'
            print(msg)
            exit()

    def log_acerto(self, msg: str):
        msg = f'{datetime.now()};{self.rpa};OK;{msg}\n'
        with open(self.caminho_log, 'a') as f:
            f.write(msg)

    def log_erro(self, msg:str):
        msg = f'{datetime.now()};{self.rpa};NOK;{msg}\n'
        with open(self.caminho_log, 'a') as f:
            f.write(msg)
    
    def log_inicio(self):
        msg = f'{datetime.now()};{self.rpa};OK;Início\n'
        with open(self.caminho_log, 'a') as f:
            f.write(msg)

    def log_fim(self):
        msg = f'{datetime.now()};{self.rpa};OK;Fim\n'
        with open(self.caminho_log, 'a') as f:
            f.write(msg)