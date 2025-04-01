import os
import json
import logging as log
import win32com.client as client
import re
from tqdm import tqdm

log.basicConfig(
    filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'Processamento_PDFC.log'),
    level=log.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Carregar configurações do JSON
with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'configPDFC.json'), 'r') as config_file:
    config = json.load(config_file)

email_field = config.get('email_field')
base_folder_path = config.get('base_folder_path')
MarckCheck = int(config.get('MarckCheck', 1))
MarckRed = int(config.get('MarckRed', 2))

try:
    outlook = client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    inbox = namespace.Folders[email_field].Folders['WorkOn']

    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    try:
        for msg in tqdm(messages, desc='Processando e-mails', unit=' e-mail'):
            sender = msg.SenderEmailAddress

            if msg.Class == 43 and msg.FlagStatus == 0 and sender == 'NoReply.Workon@bosch.com':
                try:
                    subject = msg.Subject
                    body = msg.Body
                    sender = msg.SenderName

                    standard_subject = r'Ação Requerida \[Substituted\]\s+(.*?)\s+- CR_61_Cadastro Novo Item'
                    match_subject = re.search(standard_subject, subject)

                    standard_body = r'Descrição\s+([\s\S]*?)\s+Iniciado por'
                    match_body = re.search(standard_body, body)

                    if match_subject:
                        result_subject = match_subject.group(1).strip()
                        print(f"\nTexto capturado do assunto: {result_subject}")
                    else:
                        log.info(f'Padrão de assunto do e-mail não encontrado')

                    if match_body:
                        result_body = match_body.group(1).strip()

                        standard_result = r":\s*([^|]+)\s*(?:\|\||$)"
                        values = [value.strip() for value in re.findall(standard_result, result_body)]

                        for i, value in enumerate(values, 1):
                            print(f"Valor {i}: {value}")
                    else:
                        log.info(f'Padrão do corpo do e-mail {subject} não encontrado')

                except Exception as e:
                    log.error(f'Erro ao processar e-mail: {e}', exc_info=True)

    except Exception as e:
        log.error(f'Erro no processamento dos e-mails: {e}', exc_info=True)

except Exception as e:
    log.error(f'Erro geral: {e}', exc_info=True)
