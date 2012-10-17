#!/usr/bin/env python3
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

import ezodf


PRINCIPAL = []

def carregar_registros(nome_arq):
    arq_planilha = ezodf.opendoc(nome_arq)
    planilha = arq_planilha.sheets[0]

    # Le linha de títulos
    nomes = []
    for cell in planilha.row(0):
        texto = cell.value
        if texto:
            texto = texto.strip()
            if texto[-1:] == "*":
                texto = texto[:-1]
                PRINCIPAL.append(texto)
            nomes.append(texto)

    # Separa cada um dos registros colocando os nome certos para cada célula
    registros = []
    num = planilha.nrows()
    for indice in range(1,num):
        linha = planilha.row(indice)
        if linha[0].value != None:
            registro = {}
            i = 0
            for cell in linha:
                if len(nomes) > i:
                    texto = cell.plaintext()
                    #texto = texto.replace("$","\$")
                    registro[nomes[i]] = texto
                    i += 1
            registros.append(registro)
    return registros



registros = carregar_registros("dados.ods")

DIRETORIO_GERADOS = "gerados"
if not os.path.exists(DIRETORIO_GERADOS):
    os.makedirs(DIRETORIO_GERADOS)

for registro in registros:
    modelos = registro["Modelos"].split(',')

    # Cria pasta para colocar arquivos
    if len(modelos):
        diretorio = ""
        for palavra in PRINCIPAL:
            diretorio += registro[palavra]
        diretorio = join(DIRETORIO_GERADOS, diretorio)
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
                texto = str(texto,"utf-8")
                # Faz as trocas das variáveis pelos valores corretos
                for variavel,valor in registro.items():
                    texto = texto.replace("{%s}" % variavel, "%s" % valor)
            compactado_saida.writestr(item, texto)
        compactado_modelo.close()
        compactado_saida.close()



#arq = open("modelo1.tex","r")
#modelo = arq.read()
#arq.close()
#reg = registros[0]
#
#for variavel,valor in reg.items():
#    modelo = modelo.replace(" {%s}" % variavel, " %s" % valor)
#
#arq = open("saida.tex","w")
#arq.write(modelo)
#arq.close()
