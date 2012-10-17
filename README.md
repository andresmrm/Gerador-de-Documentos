Gerador-de-Documentos
=====================

Lê uma tabela de dados ODS e, seguindo modelos ODT, vai criando documentos no mesmo formato.


Dependências
------------

- python3
- ezodf
- lxml


Como usar
---------

Planilha com dados:
- Crie uma planilha no formato ODS (padrão comum do LibreOffice) chamada **dados.ods**
- Na primeira linha da planilha coloque o nome de cada variável.
- Nas linhas seguintes coloque os valores das variáveis para cada um do registro, sendo que cada linha é um registro.
- Uma das variáveis deve ser a palavra *Modelos*. Nela coloque o nome dos modelos que devem ser aplicados a cada registro para gerar os arquivos finais.
- Pelo menos uma das vairáveis deve ter um asterisco (** \* **) no final do nome, esse será o nome da pasta onde os documentos de cada resgistro serão colocados.

Modelos:
- Dentro da pasta **modelos** crie arquivo no formado ODT (padrão comum do LibreOffice).
- Nos lugares onde quiser que entre uma variável, escreva o nome dela entre chaves.

Fim:
- Rode o programa **mimir.py** na mesma pasta em que está a planilha **dados.ods**
Os documentos serão colocados na pasta chamada **gerados**


Exemplo
-------

Caso na sua panilha ODS uma das colunas se chama **Nome** nos modelos ODT use essa variável escrevendo **{Nome}** Assim, para cada linha da sua planilha, o programa criará documentos substituindo **{Nome}** pelo valor na planilha para essa coluna e essa linha.
Examinando os arquivos de exemplo que estão nesse repositório deve ficar mais fácil de entender...
