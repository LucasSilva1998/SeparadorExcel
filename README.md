# 📊 Excel Splitter por Unidade Organizacional

Este projeto é um **console em C#** que lê uma planilha Excel e separa os dados em diferentes abas, de acordo com a coluna **"Unidade organizacional"**.  
O resultado é um novo arquivo Excel com uma aba para cada unidade distinta encontrada.

---

## 🚀 Funcionalidades
- Lê um arquivo Excel existente (`dados.xlsx`).
- Identifica automaticamente a coluna **"Unidade organizacional"**.
- Cria uma aba para cada unidade distinta.
- Copia o cabeçalho e todas as linhas correspondentes para a aba correta.
- Garante nomes válidos de abas (máx. 31 caracteres, sem caracteres inválidos).
- Evita duplicação de nomes de abas adicionando sufixos numéricos quando necessário.
- Salva o resultado em um novo arquivo (`dados_separados.xlsx`).

---

## 🛠️ Tecnologias utilizadas
- **C#** (.NET)
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) – biblioteca para manipulação de arquivos Excel (XLSX).

---

## 📂 Estrutura do projeto
Program.cs # Código principal do console

---

## ⚙️ Como executar

1. Clone este repositório:<br>
   git clone https://github.com/seu-usuario/seu-repo.git

2. Instale a biblioteca ClosedXML via NuGet:<br>
   dotnet add package ClosedXML

3. Ajuste os caminhos dos arquivos no código:<br>
  string arquivoOriginal = @"C:\Fiotec\dados.xlsx";
  string arquivoNovo = @"C:\Fiotec\dados_separados.xlsx";

4. Compile e execute:<br>
   dotnet run

O programa irá gerar um novo arquivo Excel com uma aba para cada unidade organizacional.
<br><br>

## 📌 Exemplo de uso
Suponha que o arquivo dados.xlsx tenha a seguinte estrutura:

| Nome  | Cargo      | Unidade organizacional |
|-------|------------|-------------------------|
| Ana   | Analista   | RH                      |
| João  | Gerente    | Financeiro              |
| Maria | Assistente | RH                      |

O programa irá gerar dados_separados.xlsx com duas abas:<br>

RH → contendo Ana e Maria<br>

Financeiro → contendo João<br>


## 📄 Licença
Este projeto está sob a licença MIT. Sinta-se livre para usar, modificar e distribuir.
<br>

## 🤝 Contribuições
Contribuições são bem-vindas! Abra uma issue ou envie um pull request para melhorias.
