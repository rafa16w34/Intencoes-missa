# Intenções de Missa

Este repositório contém um sistema em **Python** desenvolvido por mim (Rafael Alves Faria), para
auxiliar na marcação, organização e impressão das intenções de Missa na
**Catedral do Divino Espírito Santo -- Divinópolis/MG**.

O programa permite: 
- Cadastro de intenções para todas as categorias;
- Controle de valores pagos e recebidos (Dinheiro ou Pix);
- Cria um arquivo txt com o nome do pagador e do valor total do pix;
- Edição e exclusão de intenções já marcadas;
- Importação de intenções a partir de um arquivo `.docx`
- Exportação automática para uma planilha Excel formatada (`Intenções.xlsx`) -
- Organização automática das categorias

## Estrutura do Projeto

    IntencoesMissa/
    │── src/
    │   └── main.py
    │── requirements.txt (requisitos do ambiente virtual)
    │── README.md
    │── intencoes.docx    (Arquivo docx das intenções marcadas anteriormente)
    │── Lista_Pix.txt (gerado automaticamente)
    │── Intenções.xlsx (gerado automaticamente)

## Requisitos do sistema:

- Bibliotecas usadas e versão do Python necessária:

    ```
    python 3.10 ou superior
    et_xmlfile==2.0.0
    lxml==6.0.2
    openpyxl==3.1.5
    python-docx==1.2.0
    typing_extensions==4.15.0
    ```
    

- Instalação das Dependências

Para instalar tudo automaticamente:
```pip install -r requirements.txt```

## Como configurar o Ambiente Virtual (venv)

Para configurar o ambiente automaticamente, execute o script correspondente ao seu sistema operacional.

### Linux / macOS:
```
./setup.sh
```

### Windows (`setup.bat`)
```
setup.bat
```

### AVISOS IMPORTANTES:

- Para fazer a leitura do arquivo docx, o mesmo precisa estar nomeado como ```intencoes.docx```;
- Caso o arquivo docx esteja mal configurado a planilha pode ficar desestruturada;
- É importante verificar a planilha para corrigir possíveis erros, no caso de uma formatação ruim do arquivo docx original;
- A planilha deve ser reconfigurada para que a impressão ocorra corretamente (Ex: Tamanho das colunas e bordas);

### FUTURAS IMPLEMENTAÇÕES:

- Uma interface gráfica, ao invés de usar pelo terminal, para que esteja mais acessível para pessoas leigas;
