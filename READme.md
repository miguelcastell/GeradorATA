# Gerador de ATAs Automático

Este projeto consiste em um sistema desenvolvido em **Python** para a geração automatizada de ATAs a partir de dados contidos em uma planilha Excel. O script é ideal para automação de processos em escritórios contábeis ou administrativos que lidam com grande volume de documentos.

O sistema realiza a leitura de dados empresariais, como nome, CNPJ e responsável legal, e gera arquivos personalizados nos formatos **DOCX** e **PDF**. Além disso, conta com uma interface gráfica intuitiva que inclui barra de progresso e splash screen.

---

## 🚀 Funcionalidades

O sistema oferece um conjunto completo de ferramentas para facilitar a criação de documentos:

| Funcionalidade | Descrição |
| :--- | :--- |
| **Leitura de Excel** | Processamento automático de dados de clientes e empresas. |
| **Geração Dinâmica** | Substituição automática de variáveis no modelo de documento. |
| **Múltiplos Formatos** | Criação simultânea de arquivos em DOCX e PDF. |
| **Interface Gráfica** | Desenvolvida em Tkinter para facilitar o uso por usuários leigos. |
| **Feedback Visual** | Barra de progresso em tempo real e Splash Screen com logo. |
| **Automação de Pastas** | Abertura automática do diretório de destino ao finalizar o processo. |

---

## 📂 Estrutura do Projeto

A organização dos arquivos segue o padrão abaixo para garantir o funcionamento correto do script:

```text
gerador_ata/
│
├── gerar_atas.py          # Script principal do sistema
├── modelo_ata.docx        # Modelo base para a geração
├── lista_cliente.xlsx     # Planilha com os dados de entrada
├── logo.png               # Imagem para a splash screen
│
├── ATAS_GERADAS/          # Pasta raiz de saída
│   ├── DOCX/              # Subpasta para arquivos Word
│   └── PDF/               # Subpasta para arquivos PDF
│
└── README.md              # Documentação do projeto
```

---

## 🛠 Tecnologias Utilizadas

Para o desenvolvimento deste projeto, foram utilizadas as seguintes bibliotecas e tecnologias:

*   **Python 3**: Linguagem base do projeto.
*   **Pandas**: Manipulação e leitura de dados da planilha Excel.
*   **Python-docx**: Criação e edição de arquivos .docx.
*   **Pillow**: Processamento de imagem para a interface gráfica.
*   **Tkinter**: Biblioteca padrão para criação da interface de usuário.
*   **Docx2pdf**: Conversão automatizada de documentos Word para PDF.

---

## ⚙️ Instalação

Para configurar o ambiente e executar o projeto, instale as dependências necessárias (requirements.txt) utilizando o gerenciador de pacotes `pip`:

```bash
pip install pandas python-docx pillow docx2pdf
```

---

## ▶️ Como Usar

Siga os passos abaixo para gerar suas ATAs:

1.  **Preparação dos Dados**: Insira as informações das empresas no arquivo `lista_cliente.xlsx`.
2.  **Ajuste do Modelo**: Certifique-se de que o arquivo `modelo_ata.docx` contenha as tags de substituição correspondentes aos cabeçalhos da planilha.
3.  **Execução**: Inicie o programa executando o comando:
    ```bash
    python gerar_atas.py
    ```

---

## 📌 Observações Importantes

> *   O nome de cada arquivo gerado será automaticamente definido como o **nome da empresa** listado na planilha.
> *   O sistema gerencia a criação das pastas `ATAS_GERADAS`, `DOCX` e `PDF` de forma automática, caso elas não existam.
> *   Ao concluir o processamento de todos os itens, o Windows Explorer será aberto diretamente na pasta de resultados.

---

## 👨‍💻 Autor

**Miguel Mantoan Castellani**  
*Projeto desenvolvido para automação de escritório contábil.*
