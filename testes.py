import os
import json
import logging as log
import win32com.client as client
from tqdm import tqdm
import pdfplumber as ppl
import easyocr as ocr
import tempfile as tf
from pdf2image import convert_from_path as p2i
import cv2
import numpy as np
import zipfile
import py7zr
import rarfile
import tarfile
import shutil

# http://rb-proxy-de.bosch.com:8080
# ;http://rb-proxy-special.bosch.com:8080/

log.basicConfig(filename=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'Processamento_PDFC.log'), level=log.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assistant', 'configPDFC.json'), 'r') as config_file:
    config = json.load(config_file)
    
email_field = config.get('email_field')
base_folder_path = config.get('base_folder_path')
MarckCheck = int(config.get('MarckCheck', 1))
MarckRed = int(config.get('MarckRed', 2))

reader = ocr.Reader(['pt','en'], model_storage_directory=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ocr'))

# Salva o arquivo na pasta temp e retorna o caminho do arquivo (temp_path)
def save_temp(attachment):
    file_extension = os.path.splitext(attachment.FileName)[1] 
    '''
    file_extension          -->     *ARMAZENA A EXTENÇÃO DO ARQUIVO ANEXO*
    attachment.FileName     -->     Nome do anexo
    os.path.splitext()      -->     Divide o nome do arquivo em nome e extenção 0 = Nome 1 = Extenção
    [1]                     -->     Selecionamos a extenção
    '''
    try:
        with tf.NamedTemporaryFile(delete=False, suffix=file_extension) as file_temp:
            '''
            Cria um arquivo temporario, não permite que seja apagado, define o sufixo como file_extension (extenção do anexo) e define os arquivos atraves da variavel file_temp
            '''
            temp_path = file_temp.name
            '''
            temp_path       =       Caminho completo do arquivo na pasta temporaria
            '''
            attachment.SaveAsFile(temp_path)
            '''
            Salva o arquivo na pasta temporaria
            '''
        return temp_path
    except Exception as e:
        log.error(f'Erro ao criar arquivo temporário {attachment.FileName}: {e}')

def save_attachment(attachment, destination_folder_year, domain, receipt_date): # RETORNA UMA LISTA DE ARQUIVOS SALVOS
    try:
        timestamp = receipt_date.strftime('%Y%m%d%H%M%S') 
        destination_folder_domain = os.path.join(destination_folder_year, domain)
        if not os.path.exists(destination_folder_domain):       # os.makedirs(destination_folder_domain, exist_ok=True)
                os.makedirs(destination_folder_domain)
                log.info(f'Pasta criada: {destination_folder_domain}')
        saved_files = []
        for index, attachment in enumerate(item.Attachments, start=1):
            file_name = f'{domain}_{timestamp}_{index}_' + attachment.FileName
            destination_attachment = os.path.join(destination_folder_domain, file_name)
            if not os.path.exists(destination_attachment):
                attachment.SaveAsFile(destination_attachment)
                log.info(f'Anexo salvo em: {destination_attachment}') 
                saved_files.append(destination_attachment)
        return saved_files
    except Exception as e:
        log.error(f'Erro ao salvar o anexo {file_name}: {e}')
        return None

def status_checkmark(item, status):
    try:
        item.MarkAsTask(status)
        item.FlagStatus = status
        item.Save()    
    except Exception as e:
        log.error(f'Erro ao marcar email {attachment.FileName}: {e}')
        return None

def image_correction(image):                                                                                    # Usado diretamente no processamento da imagem e na verificação se pdf -> img (extract_text_from_compact_file)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thresh > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotated

def extract_files(file_temp, ext):
    """ Extrai arquivos compactados para uma pasta temporária e retorna a lista de arquivos extraídos. """
    extract_path = os.path.join(tf.gettempdir(), f"extracted_{ext.lstrip('.')}")
    os.makedirs(extract_path, exist_ok=True)

    try:
        if ext == '.zip':
            with zipfile.ZipFile(file_temp, 'r') as zip_ref:
                zip_ref.extractall(path=extract_path)
                file_paths = zip_ref.namelist()
        elif ext == '.7z':
            with py7zr.SevenZipFile(file_temp, 'r') as sevenz_ref:
                sevenz_ref.extractall(path=extract_path)
                file_paths = sevenz_ref.getnames()
        elif ext == '.rar':
            with rarfile.RarFile(file_temp, 'r') as rar_ref:
                rar_ref.extractall(path=extract_path)
                file_paths = rar_ref.namelist()
        # TESTAR - VERIFICAR SE ESTA FUNCIONAL
        '''
        elif ext in ('.tar', '.tar.gz', '.tgz', '.gz'):
            if ext == '.gz' and not file_temp.endswith('.tar.gz'):
                # Se for um .gz puro, descomprime sem extrair arquivos
                with gzip.open(file_temp, 'rb') as gz_ref:
                    output_path = os.path.join(extract_path, os.path.basename(file_temp).replace('.gz', ''))
                    with open(output_path, 'wb') as out_f:
                        shutil.copyfileobj(gz_ref, out_f)
                    file_paths = [output_path]
            else:
                # Extrai arquivos de .tar e .tar.gz
                with tarfile.open(file_temp, 'r:*') as tar_ref:
                    tar_ref.extractall(path=extract_path)
                    file_paths = tar_ref.getnames()
        '''
        for root, _, files in os.walk(extract_path):
            for file in files:
                src = os.path.join(root, file)
                dest = os.path.join(extract_path, file)
                if src != dest:
                    shutil.move(src, dest)

        return extract_path, file_paths
    
    except Exception as e:
        log.error(f"Erro ao extrair {ext}: {e}")
        return None, []

def process_pdfs_compressed(extract_path, poppler_path='poppler-24.08.0\\Library\\bin', dpi=300):
    """ Processa arquivos PDF extraídos e extrai o texto. """
    text_content = []
    for file in os.listdir(extract_path):
        if file.lower().endswith('.pdf'):
            file_path = os.path.join(extract_path, file)
            try:
                with ppl.open(file_path) as pdf:
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text and text.strip():
                            text_content.append(text)
                if not text_content:
                    pdf_img = p2i(file_path, dpi=dpi, poppler_path=poppler_path)
                    for img_convert in pdf_img:
                        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
                        improved_img = image_correction(image)
                        result_ocr = reader.readtext(improved_img)
                        page_text = ' '.join([text[1] for text in result_ocr])
                        if page_text:
                            text_content.append(page_text)
                log.info(f'Texto extraído do PDF {file}: {text_content[:100]}...')
            except Exception as e:
                log.error(f"Erro ao processar PDF {file}: {e}")
    return text_content

def notify_unreadable_cert(sender, file_temp):
    message = outlook.CreateItem(0)
    message.To = sender
    message.Subject = 'Illegible certificate'
    message.Body = 'The attached file is illegible. \n\nPlease return a new file.'
    message.Attachments.Add(file_temp)
    message.Save()
    message.Send()
    return None    

def clean_directory(extract_path, file_paths):
    if file_paths:
        for archive in os.listdir(extract_path):
            if os.path.isfile(os.path.join(extract_path, archive)):
                os.remove(os.path.join(extract_path, archive))
            elif os.path.isdir(os.path.join(extract_path, archive)):
                shutil.rmtree(os.path.join(extract_path, archive))
    return None

try:
    outlook = client.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace('MAPI')
    try:
        inbox = namespace.Folders[email_field].Folders['TESTE PDFC']      # ALTERAR  - inbox = namespace.GetDefaultFolder(6)
    except Exception as e:                                                  # REMOVER #
        try:                                                                # ------- #       
            inbox = namespace.Folders[email_field].Folders['Inbox']         # ------- #
        except Exception as e:                                              # ------- #
            log.error(f'{e}')                                               # ------- #
    try:
        for item in tqdm(inbox.Items, desc='Processando e-mails', unit=' e-mail'):
            switch = None
            if item.Class == 43 and item.FlagStatus == 0:
                try:
                    sender = item.SenderEmailAddress
                    sender = 'pedro@bosch.com'
                    domain = sender.split('@')[1].split('.')[0] 
                    receipt_date = item.ReceivedTime
                    receipt_year = receipt_date.year
                    
                    if domain and receipt_year:
                        try:
                            destination_folder_year = os.path.join(base_folder_path, f'Arquivo {receipt_year}')
                            if not os.path.exists(destination_folder_year):
                                os.makedirs(destination_folder_year)
                                log.info(f'Pasta criada: {destination_folder_year}')
                        except Exception as e:                                             
                                log.error(f'{e}')

                        for attachment in item.Attachments:
                            try:
                                if attachment.FileName.lower().endswith(('.pdf')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        with ppl.open(file_temp) as pdf:
                                            for page in pdf.pages:
                                                text = page.extract_text()
                                                if text and text.strip():
                                                    '''
                                                    if True:
                                                        notify_unreadable_cert(sender, file_temp)
                                                    '''
                                                    
                                                    '''
                                                    LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                    '''
                                                    #save_attachment(file_temp)
                                                    log.info(f'Texto extraído do arquivo pdf {attachment.FileName}: {text[:50]}...')
                                                    switch = 1
                                                if not text:
                                                    pdf_img = p2i(file_temp, dpi=300, poppler_path='poppler-24.08.0\\Library\\bin')
                                                    for i, img_convert in enumerate(pdf_img):
                                                        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
                                                        improved_img = image_correction(image)
                                                        result_ocr = reader.readtext(improved_img)
                                                    '''
                                                    NOTIFICAR COM ARQUIVO RUIM
                                                    if True:
                                                        notify_unreadable_cert(sender, file_temp)
                                                    '''
                                                    '''
                                                    Else:
                                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                    '''
                                                    #save_attachment(file_temp)
                                                    log.info(f'Texto extraído da img {attachment.FileName}: {result_ocr[:50]}...')
                                                    switch = 1
                                    os.remove(file_temp)
                            except Exception as e:
                                log.error(f'{attachment.FileName}: {e}')  
                                switch = 0
                            
                            try:    
                                if attachment.FileName.lower().endswith(('.jpg', '.png')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        result_ocr = reader.readtext(file_temp)
                                        if result_ocr:
                                            '''
                                            if True:
                                                notify_unreadable_cert(sender, file_temp)
                                            '''
                                            '''
                                            else:
                                                LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                USAR result_ocr (ONDE O TEXTO É RETORNADO)                      
                                            '''
                                            # save_attachment(file_temp)
                                            log.info(f'Texto extraído do arquivo de imagem {attachment.FileName}: {result_ocr[:50]}...')
                                            switch = 1
                                    os.remove(file_temp)
                            except Exception as e:
                                log.error(f'{attachment.FileName}: {e}')
                                switch = 0
                            
                            try:       
                                if attachment.FileName.lower().endswith(('.zip', '.7z', '.rar', '.tar', '.gz')):
                                    file_temp = save_temp(attachment)
                                    if file_temp:
                                        
                                        ext = os.path.splitext(file_temp)[1].lower()
                                        extract_path, file_paths = extract_files(file_temp, ext)

                                        '''
                                        extract_files(file_temp, ext) === PAREI AQUI

                                        '''

                                        if extract_path:
                                            for file in os.listdir(extract_path):
                                                file_path = os.path.join(extract_path, file)
                                                if file.lower().endswith(('.pdf')):
                                                    try:
                                                        with ppl.open(file_path) as pdf:
                                                            for page in pdf.pages:
                                                                text = page.extract_tables()
                                                                if text and text.strip():
                                                                    '''
                                                                    if True:
                                                                        notify_unreadable_cert(sender, file_temp)
                                                                    '''
                                                                    '''
                                                                    else:
                                                                        LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                    '''
                                                                    # save_attachment(file_path)
                                                                    log.info(f'Texto extraído do arquivo pdf {file}: {text[:50]}...')
                                                                    switch = 1
                                                    except Exception as e:
                                                        log.error(f"{file}: {e}")
                                                        switch = 0

                                                    if not text:
                                                        try:    
                                                            with ppl.open(file_path) as pdf:
                                                                for page_num, _ in enumerate(pdf.pages):
                                                                    pdf_img = p2i(file_path, dpi=150, first_page=page_num+1, last_page=page_num+1, poppler_path='poppler-24.08.0\\Library\\bin')
                                                                    
                                                                    for img_convert in pdf_img:
                                                                        image = cv2.cvtColor(np.array(img_convert), cv2.COLOR_RGB2BGR)
                                                                        improved_img = image_correction(image)
                                                                        result_ocr = reader.readtext(improved_img)
                                                                        
                                                                        extracted_text = " ".join(text[1] for text in result_ocr)
                                                                        
                                                                        if extracted_text.strip():  
                                                                            '''
                                                                            if True:
                                                                                notify_unreadable_cert(sender, file_temp)
                                                                            '''
                                                                            '''
                                                                            else:
                                                                                LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                                            '''
                                                                            # save_attachment(file_path)
                                                                            log.info(f'Texto extraído do arquivo {file_path}: {extracted_text[:50]}...')
                                                                            switch = 1
                                                        except Exception as e:
                                                            log.error(f"{file_path}: {e}")
                                                            switch = 0
                                                    '''
                                                    ACRESCENTAR MELHORIA DE IMAGEM
                                                    '''
                                                elif file.lower().endswith(('.jpg', '.png')):
                                                    try:
                                                        result_ocr = reader.readtext(file_path)
                                                        if result_ocr:
                                                            '''
                                                            if True:
                                                                notify_unreadable_cert(sender, file_temp)
                                                            '''
                                                            '''
                                                            else:
                                                                LOGICA PARA IDENTIFICAÇÃO SE O CONTEUDO É CERTIFICADO
                                                            '''
                                                                # save_attachment(file_path)
                                                            log.info(f'Texto extraído do arquivo de imagem {file}: {text[:50]}...')
                                                            switch = 1
                                                    except Exception as e:
                                                        log.error(f"{file}: {e}")
                                                        switch = 0
                                            
                                            clean_directory(extract_path, file_paths)
                                        os.remove(file_temp)
                            
                            except Exception as e:
                                log.error(f'{attachment.FileName}: {e}')

                except Exception as e:
                    log.error(f'{sender}: {e}') 

            if switch == 1:
                status_checkmark(item, MarckCheck)
                item.Move(namespace.Folders[email_field].Folders['PDFC_Processed'])
            elif switch == 0:
                item.MarkAsTask(MarckRed)
                item.FlagStatus = MarckRed
                item.Save()
                item.Move(namespace.Folders[email_field].Folders['PDFC_Alert']) 

    except Exception as e:
        log.error(f'{e}')
    
except Exception as e:
    log.error(f'{e}')
