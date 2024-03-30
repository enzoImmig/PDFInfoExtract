#Libs ---------------------------------
import pytesseract
from pdf2image import convert_from_path
import re
import pandas as pd

#Constants ---------------------------------
spreadsheet_path = 'testSpreadsheet.xlsx'
pdf_file = 'tarefa1.pdf'
rgxRules = {
    'titulo': "(?<=TAREFA [0-9][0-9]...)(.*)",
    'num': "(?<=TAREFA )[0-9]{0,2}",
    'instructs':"(?<=Instruções: )(.*);",
    'pontuacao':"(?<=Pontuação: )[0-9]{0,4}",
    'entrega':"(?<=Entrega: Lote )(.*);",
    'local':"(?<=Local: n. )(.*);",
    'autor':"(?<=Boa sorte a todas as equipes! )(.*)\.",
}

#Globals ----------------------------------
pages = convert_from_path(pdf_file)
dictInfo = dict()
infoNames = []

def init_Config(arrNames):
    for name in arrNames:
        dictInfo[name]=""
        infoNames.append(name)

def extract_text_from_image(image):
    text = pytesseract.image_to_string(image, lang="por")
    return text

def run():
    #extrai as informações setadas em infoNames a adiciona ela no dicionario
    for page in pages:
        text = extract_text_from_image(page)
        
        #aplica os regexs
        for info in infoNames:
            dictInfo[info] = re.findall(rgxRules[info], text)[0]
        
    dtTasks = pd.read_excel(spreadsheet_path)

    dataRow = []
    for info in infoNames:
        dataRow.append(dictInfo[info])

    dtNewTask = pd.DataFrame(
        data=[dataRow],
        columns=infoNames,
    )

    dtUpdated = pd.concat([dtTasks, dtNewTask])
    dtUpdated.to_excel("testSpreadsheet.xlsx", index=False)

if __name__ == '__main__':
    init_Config(rgxRules.keys())
    run()