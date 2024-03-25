import shutil
import tempfile
import os 
from smb.SMBConnection import SMBConnection
from datetime import datetime, timedelta
import pandas as pd
import xmltodict
from unidecode import unidecode
from bs4 import BeautifulSoup

def setPathDate():
    #pega a data atual
    data_atual = datetime.now()
    #ajusta para filtrar e salvar a pasta com a data de domingo
    data_inicio = datetime(data_atual.year,data_atual.month,data_atual.day-1).strftime('%Y%m%d')
    dir_base = 'D:/'+data_inicio
    dir_inicio = 'D:/'
    #verifica se a pasta já existe
    if not os. path. isdir(dir_base):
        os.mkdir(dir_base)
    return dir_base, dir_inicio, data_inicio

def conectAndMOveFiles(dir_base, data_inicio):
    #define os dados para conectar no master
    server_ip = "172.0.0.0" # Coloque aqui o IP do seu servidor
    server_name = 'myserver' # Coloque aqui o nome do seu servidor
    share_name = "preco$" # Esse é o nome do principal diretório da rede onde você deseja conectar
    network_username = 'usuario' # Esse é o seu nome de usuário na rede
    network_password = 'senha' # Essa é sua senha de rede
    machine_name = 'mypc' # O nome do seu computador. Para saber, acesse o terminal
    #conecta na pasta compartilhada no master
    conn = SMBConnection(network_username, network_password, machine_name, server_name, use_ntlm_v2 = True)
    assert conn.connect(server_ip)
    #pega todas as pastas contidas dentro da pasta gerada no domingo
    files = conn.listPath(share_name, data_inicio+"/") # Coloque aqui o caminho completo
    #passa por todas as subpastas geradas
    for item in files:
        if len(item.filename) > 3: #pega apenas o que for pasta real
            new_path = data_inicio+"/"+item.filename+ "/"
            files_p = conn.listPath(share_name, new_path)
            for item_p in files_p:#passa por todos os arquivos da pasta
                if item_p.filename[0:8] in ['00601002','00601051']: #filtra apenas o que for de CD
                    file_path = new_path+item_p.filename
                    sf = conn.getAttributes(share_name, new_path+item_p.filename)
                    #gera o arquivo temporario para mover
                    file_obj = tempfile.NamedTemporaryFile(mode='w+b', delete=False)
                    file_name = file_obj.name
                    #pega os dados do arquivo
                    file_attributes, copysize = conn.retrieveFile(share_name, file_path, file_obj)
                    file_obj.close()
                    #verifica se o aquivo já existe e se não existir copia para a pasta do drive
                    if not os.path.exists(dir_base +"/"+item_p.filename):
                        shutil.copy(file_name, dir_base +"/"+item_p.filename)
    conn.close()

def alternativeXmltoExcel(files, data, dir_base):
    for itFile in files:
        xml_data = open(dir_base+"/"+itFile, 'r').read() 
        soup = BeautifulSoup(xml_data, "html.parser")
        franquia = itFile.replace('00601002','').replace('00601051','').replace(' - ','')[0:8]
        #xmlDict = xmltodict.parse(xml_data)
        all_data = []
        for obs in soup.select("Row"):
            lineData = []
            lineData.append(franquia)
            for dat_ in obs.find_all('data'):                
                lineData.append(dat_.get_text())
                if len(lineData) >= 3 and len(lineData[2]) == 0:                    
                    break
            if len(lineData) == 21:
                data.append(lineData)
    return data

def fixFilesXmlExcel(dir_base, dir_ini):
    cols = ['franquia']
    data = []
    lAcols = False
    fileNotProcess = []
    for _, _, arquivo in os.walk(dir_base):
        for file in arquivo:
            try:
                xml_data = open(dir_base+"/"+file, 'r').read() 
                xmlDict = xmltodict.parse(xml_data)
                franquia = file.replace('00601002','').replace('00601051','').replace(' - ','')[0:8]
                #print(franquia)
                #print(xmlDict['Workbook']['Worksheet']['Table']['Row'])
                data_ = xmlDict['Workbook']['Worksheet']['Table']['Row']
                count = 0
                for line in data_:
                    #print(line['Cell']['Data']['#text'])
                    if 'Cell' in line and type(line['Cell']) == type([]):
                        lineData = []
                        count = 0
                        lineData.append(franquia)
                        for subline in line['Cell']:
                            if '#text' in subline['Data']:
                                if not lAcols and subline['@ss:StyleID'] == 'HEADER':
                                    coluna = subline['Data']['#text'].replace('í','i').replace('%','').replace('ç','c').replace('ã','a')
                                    cols.append(coluna)
                                if subline['@ss:StyleID'] == 'LINEALeftGeneral':
                                    count += 1
                                    lineData.append(subline['Data']['#text'])
                               # if count == 9:
                                    #print(subline)
                                    #print("tamanho da line", len(line['Cell']))
                            else:
                                lineData.append('-')
                            #print(subline)
                        #print("tamanho da line", len(line['Cell']))
                        #print("tamanho do arr", count)
                        if len(lineData) > 3 and lineData[4] != '-':
                            data.append(lineData)
                    if len(cols) > 5:
                        lAcols = True
            except Exception as e:
                fileNotProcess.append(file)
                #print("Erro ao ler XML", repr(e), file)

            
            
        break
    
    data =  alternativeXmltoExcel(fileNotProcess, data, dir_base)
    #print(data)
    df = pd.DataFrame(data)  # Create DataFrame and transpose it.
    #print(cols)
    oldcols = cols
    cols = []
    for coluna in oldcols:
        cols.append(unidecode(coluna).replace('A(c)','e').replace('ASS','c').replace('APS','a').replace('Aq','iq'))
    df.columns = cols
    df.to_excel(dir_ini+"/hoje.xlsx", index = False)


diretorio, diretorio_inicio, data_inicio = setPathDate()
conectAndMOveFiles(diretorio, data_inicio)
fixFilesXmlExcel(diretorio, diretorio_inicio)