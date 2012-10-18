#!/usr/bin/env python
# -*- coding: utf-8 -*-

#-----------------------------------------------------------------------------
# Copyright 2012 Andrés M. R. Martano
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>
#-----------------------------------------------------------------------------

import os
from os.path import join
import zipfile

import uno


def carregar_registros(oSheet):
    lCursor = oSheet.createCursor()
    lCursor.gotoStartOfUsedArea(False)
    lCursor.gotoEndOfUsedArea(True)
    nLinhas = lCursor.getRows().getCount()
    nColunas = lCursor.getColumns().getCount()
    # Le linha de títulos
    nomes = []
    chaves_primarias = []
    for i in xrange(0,nColunas):
        texto = oSheet.getCellByPosition(i,0).getString()
        if texto:
            texto = texto.strip()
            if texto[-1:] == "*":
                texto = texto[:-1]
                chaves_primarias.append(texto)
            nomes.append(texto)
    # Separa cada um dos registros colocando os nomes certos para cada célula
    registros = []
    for j in xrange(1,nLinhas):
        i = 0
        texto = oSheet.getCellByPosition(i,j).getString()
        if texto:
            registro = {}
            for i in xrange(0,nColunas):
                if len(nomes) > i:
                    texto = oSheet.getCellByPosition(i,j).getString()
                    registro[nomes[i]] = texto
            registros.append(registro)
    return registros, chaves_primarias

def obter_documento():
    context = uno.getComponentContext()
    smgr = context.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop",context)
    document = desktop.getCurrentComponent()
    return document

# Altera o diretorio de trabalho para o diretorio do documento
def alterar_diretorio(document):
    url = document.URL
    syspath = uno.fileUrlToSystemPath(url)
    directory = os.path.dirname(syspath)
    os.chdir(directory)

# Gera os arquivos baseado nos registros passados
def gerar_arquivos(registros, chaves_primarias, nome_pasta):
    if not os.path.exists(nome_pasta):
        os.makedirs(nome_pasta)
    for registro in registros:
        modelos = registro["Modelos"].split(',')
        # Cria pasta para colocar arquivos
        if len(modelos):
            diretorio = ""
            for palavra in chaves_primarias:
                diretorio += registro[palavra]
            diretorio = join(nome_pasta, diretorio)
            if not os.path.exists(diretorio):
                os.makedirs(diretorio)
        # Cria um arquivo para cada modelo pedido
        for modelo in modelos:
            modelo = modelo.strip()
            caminho_modelo = join("modelos","%s.odt" % modelo)
            compactado_modelo = zipfile.ZipFile(caminho_modelo,"r")
            caminho_saida = join(diretorio,"%s.odt" % modelo)
            compactado_saida = zipfile.ZipFile(caminho_saida,"w")
            for item in compactado_modelo.infolist():
                texto = compactado_modelo.read(item.filename)
                if item.filename == "content.xml":
                    texto = texto.decode('utf-8')
                    # Faz as trocas das variáveis pelos valores corretos
                    for variavel,valor in registro.items():
                        texto = texto.replace("{%s}" % variavel, "%s" % valor)
                    texto = texto.encode('utf-8')
                compactado_saida.writestr(item, texto)
            compactado_modelo.close()
            compactado_saida.close()


def PRINCIPAL():
    document = obter_documento()
    alterar_diretorio(document)
    nPlanilhas = document.getSheets().getCount()
    for i in range(0, nPlanilhas):
        oSheet = document.getSheets().getByIndex(i)
        registros, chaves_primarias = carregar_registros(oSheet)
        nome_planilha = oSheet.getName()
        gerar_arquivos(registros, chaves_primarias, nome_planilha)

