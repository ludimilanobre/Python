# -*- coding: utf-8 -*-
"""
Created on Tue Mar  9 20:13:04 2021

@author: ludim
"""

## Com a ajuda de alguns fóruns consegui montar esse breve código que converte arquivos .docx em arquivos .pdf. 
## Para esse caso os arquivos em .pdf foram salvos em uma pasta diferente do diretório.

import os
import win32com.client as win32

## Atualizei o diretório para o caminho a seguir:
## in_file = 'D:\CEBRASPE\CelpeBras\Ano_2021\Certificados\PARTE_ESCRITA\AVALIADOR\Word'

wdFormatPDF = 17

def convertFile(file):
    
    ## Pasta em que quero salvar os arquivos em PDF:
    out_path = 'D:\CEBRASPE\CelpeBras\Ano_2021\Certificados\PARTE_ESCRITA\AVALIADOR\PDF'
    
    ## A função 'path.abspath' do módulo 'os' retorna o caminho efetivamente utilizado para salvar os arquivos.
    in_file = os.path.abspath(file)
    out_file = os.path.abspath(out_path + '\\' + file.replace(".docx", ".pdf"))
    
    word = win32.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat = wdFormatPDF)
    doc.Close()
    word.Quit()


for file in os.listdir("."):
    if file.endswith(".doc") or file.endswith(".docx"):
        convertFile(file)
        