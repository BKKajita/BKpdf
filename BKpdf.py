# BKpdf manipulator by Bruno K. Kajita
from pathlib import Path
import PySimpleGUI as sg
import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger

def pdfRotate(path, rotateFileName, filesDirectory, rotationValue):
    pdf = PdfReader(path)
    pdf_writer = PdfWriter()
    for i in range(len(pdf.pages)):
        page = pdf.getPage(i)
        pdf_writer.addPage(page)
        pdf_writer.pages[i].rotate(int(rotationValue))
    output_filename = '{}/{}.pdf'.format(filesDirectory, rotateFileName)
    with open(output_filename, 'wb') as file:
        pdf_writer.write(file)
    sg.popup("Rotação concluída!")

# Mescla páginas de diferentes arquivos PDF
def merge_pdf(selected_pdf_files, filesDirectory, pdf_file_name):
   
    # separa o nome de cada arquivo pdf usando ; como separador
    pdf_files = selected_pdf_files.split(";")
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    path_to_file = "{}/{}.pdf".format(filesDirectory, pdf_file_name)
    merger.write(path_to_file)
    merger.close()
    sg.popup("Páginas mescladas com sucesso!")

 
# Separa cada página de um arquivo PDF com diversas páginas em arquivos individuais
def split_pdf(path, pdfName, filesDirectory):
    pdf = PdfReader(path)
    for page in range(len(pdf.pages)):
        pdf_writer = PdfWriter()
        pdf_writer.add_page(pdf.pages[page])
        output_filename = '{}/{}_pg_{}.pdf'.format(filesDirectory,
            pdfName, page+1)
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
    sg.popup("Separação concluída!")

 
# Protege o PDF contra impressões
def noPrint_pdf(path, npPdf, filesDirectory):
    pdf = PdfReader(path)
    pdf_writer = PdfWriter()
    for i in range(len(pdf.pages)):
        page = pdf.getPage(i)
        pdf_writer.addPage(page)
    pdf_writer.encrypt(user_pwd='', owner_pwd=None, use_128bit=True, permissions_flag=10)
    output_filename = '{}/{}.pdf'.format(filesDirectory, npPdf)
    with open(output_filename, 'wb') as file:
        pdf_writer.write(file)
    sg.popup("Bloqueio concluído!")

# selected_pdf_files, filesDirectory, pdf_file_name
def mergePdf():
    mergePdf_layout = [
        [sg.Text("Escolha uma das opções a seguir:")],
        [sg.Text("Arquivos"), sg.Input(), sg.FilesBrowse("Procurar", file_types=(("Arquivos PDF","*.pdf"),), key="-FILES_TO_MERGE-")],
        [sg.Text("Informe um nome para o arquivo mesclado: "), sg.InputText(key="-MERGED_FILE_NAME-")],
        [sg.Text("Destino"), sg.Input(), sg.FolderBrowse("Escolher", key="-MERGED_FILE_LOCATION-")],
        [sg.Ok(), sg.Cancel()]
    ]
    
    merge_pdf_window = sg.Window("Rotate PDF", mergePdf_layout)

    while True:
        event, values = merge_pdf_window.read()
        print(event, values)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "Ok":
            merge_pdf(
                selected_pdf_files = values["-FILES_TO_MERGE-"],
                pdf_file_name = values["-MERGED_FILE_NAME-"],
                filesDirectory = values["-MERGED_FILE_LOCATION-"],
            )
            break

    merge_pdf_window.close()  

def splitPdf():
    splitPdf_layout = [
        [sg.Text("Arquivos"), sg.Input(), sg.FileBrowse("Procurar", file_types=(("Arquivos PDF","*.pdf"),), key="-FILE_TO_SPLIT-")],
        [sg.Text("Informe um nome para os arquivos: "), sg.InputText(key="-SPLITTED_FILES-")],
        [sg.Text("Destino"), sg.Input(), sg.FolderBrowse("Escolher", key="-SPLIT_FILES_LOCATION-")],
        [sg.Ok(), sg.Cancel()]
    ]
    
    split_pdf_window = sg.Window("Rotate PDF", splitPdf_layout)

    while True:
        event, values = split_pdf_window.read()
        print(event, values)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "Ok":
            split_pdf(
                path = values["-FILE_TO_SPLIT-"],
                pdfName = values["-SPLITTED_FILES-"],
                filesDirectory = values["-SPLIT_FILES_LOCATION-"],
            )
            break

    split_pdf_window.close() 

def noPrintPdf():
    noPrintPdf_layout = [
        [sg.Text("Arquivos"), sg.Input(), sg.FileBrowse("Procurar", file_types=(("Arquivos PDF","*.pdf"),), key="-FILE_TO_PROT-")],
        [sg.Text("Informe um nome para o arquivo: "), sg.InputText(key="-PROTECTED_FILE_NAME-")],
        [sg.Text("Destino"), sg.Input(), sg.FolderBrowse("Escolher", key="-PROTECTED_FILE_LOCATION-")],
        [sg.Ok(), sg.Cancel()]
    ]
    
    no_print_pdf_window = sg.Window("Rotate PDF", noPrintPdf_layout)

    while True:
        event, values = no_print_pdf_window.read()
        print(event, values)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "Ok":
            noPrint_pdf(
                path = values["-FILE_TO_PROT-"],
                npPdf = values["-PROTECTED_FILE_NAME-"],
                filesDirectory = values["-PROTECTED_FILE_LOCATION-"],
            )
            break

    no_print_pdf_window.close()  

def rotatePDF():
    rotate_pdf_layout = [
        [sg.Text("Escolha o grau de rotação das páginas: "), sg.Combo(['90','180','270'], key="-ANGLE-")],
        [sg.Text("Arquivo"), sg.Input(), sg.FileBrowse("Procurar", file_types=(("Arquivos PDF","*.pdf"),), key="-ORIGINAL_FILE_PATH-"),],
        [sg.Text("Nome do arquivo rotacionado:"), sg.InputText(key="-FILE_NAME-")],
        [sg.Text("Destino"), sg.Input(), sg.FolderBrowse("Escolher", key="-SAVE_LOCATION-")],
        [sg.Ok(), sg.Cancel()]
    ]
    
    rotate_pdf_window = sg.Window("Rotate PDF", rotate_pdf_layout)

    while True:
        event, values = rotate_pdf_window.read()
        print(event, values)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "Ok":
            pdfRotate(
                path = values["-ORIGINAL_FILE_PATH-"],
                rotateFileName = values["-FILE_NAME-"],
                filesDirectory = values["-SAVE_LOCATION-"],
                rotationValue = values["-ANGLE-"]
            )
            break

    rotate_pdf_window.close()

#  janela inicial
def main_window():
    menu_def = [["Informações",["Sobre","Ajuda","Histórico de versões"]]]
    layout = [
        [sg.MenubarCustom(menu_def, tearoff=False)],
        [sg.Text("Escolha uma das opções a seguir:")],
        [sg.Button("Mesclar", size=(15,2)), sg.Button("Separar", size=(15,2))],
        [sg.Button("Bloquear Impressão", size=(15,2)), sg.Button("Rotacionar", size=(15,2))],
        [sg.Exit("Sair", size=(5,1),button_color=("#475841"))]
    ]

    window = sg.Window("BKpdf", layout)
 
    while True:
        event, values = window.read()
        print(event, values)
        if event in (sg.WINDOW_CLOSED, "Sair"):
            break
        if event == "Sobre":
            window.disappear()
            sg.popup("BK PDF Manipulator\nVersão 3.0.0\nDesenvolvido por Bruno K. Kajita")
            window.reappear()
        if event == "Ajuda":
            window.disappear()
            sg.popup("Escolhar 'Mesclar' para unir diferentes arquivos PDF em um único.\n Escolha 'Separar' para criar um arquivo para cada página.")
            window.reappear()
        if event == "Histórico de versões":
            window.disappear()
            sg.popup("Versão 1.0.0 - Inicial\nVersão 2.0.0 - PDF não imprimível \nVersão 2.0.1 - Correções\nVersão 3.0.3 - Divesas Melhorias")
            window.reappear()
        if event == "Mesclar":
            mergePdf()
        if event == "Separar":
            splitPdf()
        if event == "Bloquear Impressão":
            noPrintPdf()
        if event == "Rotacionar":
            rotatePDF()
           
    window.close()


if __name__ == "__main__":
    sg.theme('Topanga')
    main_window()