import pytesseract
from pdf2image import convert_from_path
import re
import pandas as pd

pdf_file = 'tarefa1.pdf'
pages = convert_from_path(pdf_file)
dictInfo = dict(titulo="", num="", autor="", pontuacao="", entrega="", local="", instructs="")

rgxRules = {
    'titulo': "(?<=TAREFA [0-9][0-9]...)(.*)",
    'num': "(?<=TAREFA )[0-9]{0,2}",
    'instructs':"(?<=Instruções: )(.*);",
    'pontuacao':"(?<=Pontuação: )[0-9]{0,4}",
    'entrega':"(?<=Entrega: Lote )(.*);",
    'local':"(?<=Local: n. )(.*);",
    'autor':"(?<=Boa sorte a todas as equipes! )(.*)\.",
}

infoNames = ["titulo", "num", "instructs", "pontuacao", "entrega", "local", "autor"]

def extract_text_from_image(image):
    text = pytesseract.image_to_string(image, lang="por")
    return text

#extrai as informações setadas em infoNames a adiciona ela no dicionario
for page in pages:
    text = extract_text_from_image(page)
    
    #aplica os regexs
    for info in infoNames:
        dictInfo[info] = re.findall(rgxRules[info], text)[0]
    
dtTasks = pd.read_excel('testSpreadsheet.xlsx')
firstFreeRowIdx = dtTasks['Tarefa'].last_valid_index()+1

dataRow = [[dictInfo["num"], dictInfo["titulo"], dictInfo["pontuacao"], dictInfo["entrega"], dictInfo["instructs"], dictInfo["autor"],""]]

dtNewTask = pd.DataFrame(
    data=dataRow,
    columns=['Tarefa', 'Título', 'Pontuação', 'Lote', 'Descrição', 'Autor','Comentários'],
    index=[firstFreeRowIdx]
)

dtUpdated = pd.concat([dtTasks, dtNewTask])
dtUpdated.to_excel("testSpreadsheet.xlsx", index=False)