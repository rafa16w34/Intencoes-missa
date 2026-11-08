# Sistema de Intenções de Missa — Catedral do Divino Espírito Santo

Sistema web desenvolvido em **Django / Python** para auxílio no atendimento de balcão, registro, conferência financeira e impressão litúrgica das intenções de Missa da **Catedral do Divino Espírito Santo (Divinópolis/MG)**.

---

## Principais Funcionalidades

- **Painel de Celebrações com Timers Coloridos**:
  - Visualização em cards sequenciais por data e horário;
  - **Verde**: Faltam mais de 24 horas;
  - **Amarelo**: Faltam poucas horas para a celebração;
  - **Vermelho**: Faltam 10 minutos para o horário marcado;
  - **Cinza**: Missa já ocorrida (bloqueio automático de novas marcações após 3 horas).
- **Menu Lateral Retrátil (Offcanvas)**:
  - Navegação prática acionada pelo botão no canto superior esquerdo.
- **Agrupamento Litúrgico com Dropdowns**:
  - Organização das intenções em abas/dropdowns (Honra e Louvor, Aniversários, Casamentos, Ação de Graças, Intenções Particulares, Recuperação de Saúde, 7º Dia e Irmãos Falecidos).
- **Atendimento por Fiel e Controle de Pagamentos**:
  - Todas as intenções de uma mesma pessoa são amarradas a um único lote de pagamento;
  - Suporte a Dinheiro, Pix (com formatação automática de inicial maiúscula) e Doações do troco;
  - **Tratamento de Pagamentos Pendentes**: Intenções suspensas não são somadas ao caixa e não são impressas na folha do celebrante até a quitação definitiva em lote.
- **Exportação Integrada (Botão "Baixar")**:
  - **Word (`.docx`)**: Folha oficial diagramada no padrão da diocese com brasões e cabeçalho;
  - **PDF (`.pdf`)**: Conversão fiel via LibreOffice, com mesma formatação e pronta para impressão;
  - **Excel (`.xlsx`)**: Planilha formatada em duas abas (Intenções em 4 colunas + Fechamento de Caixa);
  - **Texto (`.txt`)**: Espelho de auditoria paroquial do caixa (`Registro_Caixa.txt`).
- **Importação Inteligente de DOCX da Secretaria**:
  - Identifica data e horário automaticamente;
  - Reconhece o asterisco (`*`) para distinguir intenções da secretaria das intenções marcadas na sessão.

---

## Estrutura do Repositório

```text
Intencoes-missa/
│── core/                      # Configurações centrais do Django (settings, urls, wsgi)
│── intencoes/                 # App principal de intenções
│   ├── templates/intencoes/   # Telas HTML (painel, detalhes, cadastro, pagamento, caixa)
│   ├── migrations/            # Migrações do banco de dados SQLite
│   ├── docx_generator.py      # Gerador oficial da folha Word (.docx)
│   ├── pdf_generator.py       # Gerador idêntico em PDF (.pdf)
│   ├── excel_generator.py     # Gerador da planilha com abas de intenções e caixa (.xlsx)
│   ├── docx_importer.py       # Leitor e importador de intenções (.docx)
│   ├── models.py              # Modelos: Missa, IntencaoMissa, Pagamento, RegistroCaixa
│   ├── forms.py               # Formulários validados
│   └── views.py               # Lógica de controle e regras de negócio
│── static/                    # Arquivos estáticos (brasões da diocese e catedral)
│   └── img/
│── templates/                 # Template base compartilhado (base.html)
│── scripts/                   # Scripts de inicialização e bancos de teste
│   ├── iniciar_oficial.sh / .bat
│   ├── iniciar_teste.sh / .bat
│   └── zerar_banco_teste.sh / .bat
│── exemplo/                   # Arquivo docx de exemplo para importação
│── terminal_legado/           # Versão histórica original que operava via terminal (CLI)
│── iniciar.sh / iniciar.bat   # Atalhos rápidos na raiz para execução
│── db.sqlite3                 # Banco de dados SQLite oficial
│── manage.py                  # CLI do Django
│── requirements.txt           # Dependências do projeto
└── README.md
```

---

## Como Executar

### 1. Pré-requisitos
- Python 3.10 ou superior
- LibreOffice (necessário para a conversão idêntica em PDF)

### 2. Instalação das Dependências
```bash
pip install -r requirements.txt
```

### 3. Iniciar o Sistema

#### No Linux / macOS:
```bash
./iniciar.sh
```
*(ou `./scripts/iniciar_oficial.sh`)*

#### No Windows:
Basta dar dois cliques em **`iniciar.bat`** na raiz do projeto *(ou `scripts\iniciar_oficial.bat`)*.

O sistema abrirá no endereço local: **`http://127.0.0.1:8000/`**.

---

## Modo de Testes (Ambiente Seguro)
Para treinar atendentes ou realizar testes sem alterar os dados oficiais da paróquia:
- **Linux**: `./scripts/iniciar_teste.sh`
- **Windows**: `scripts\iniciar_teste.bat`
- **Para limpar o banco de testes**: `./scripts/zerar_banco_teste.sh` ou `scripts\zerar_banco_teste.bat`.

---

## Testes Automatizados
O projeto conta com suíte de testes unitários cobrindo regras financeiras, timers, lotes pendentes e exportações:
```bash
python manage.py test
```
