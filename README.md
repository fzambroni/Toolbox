# Ortems Toolbox

**Ortems Toolbox** é uma aplicação Windows criada em AutoIt para montar, revisar, validar e importar dados de demonstração em uma base **Ortems SQL Server**.

A ferramenta substitui o fluxo antigo baseado em planilha macro, oferecendo uma interface tabulada para cadastro de calendários, máquinas, operações, roteiros, itens, BOM, ordens de produção, recursos secundários, capacidade e movimentos de estoque. Depois que os dados são preenchidos, o app gera um script SQL estruturado e pode executá-lo diretamente no banco conectado.

> ⚠️ **Uso recomendado:** bases de demonstração, treinamento, sandbox ou ambientes clonados. Não execute em banco produtivo sem backup, revisão do SQL gerado e validação com o responsável pelo ambiente.

---

## Visão geral

O Toolbox foi pensado para acelerar a preparação de bases Ortems, especialmente quando é necessário criar rapidamente um cenário de planejamento para demo, treinamento, prova de conceito ou validação funcional.

Principais objetivos:

- reduzir a dependência de planilhas Excel com macros;
- facilitar a criação de dados Ortems consistentes;
- validar referências antes da importação;
- gerar SQL de forma rastreável;
- permitir importação/exportação em workbook Excel para round-trip;
- preservar as Work Orders em estado **lançável**, sem pré-criar registros de WIP/lançamento.

---

## Funcionalidades

### Conexão com SQL Server

- Conexão com banco Ortems SQL Server.
- Suporte a:
  - Windows Authentication;
  - SQL Server Authentication.
- Exibição da connection string ativa.
- Teste de conexão via `ADODB.Connection`.
- Ferramenta **Inspect Table** para visualizar colunas de uma tabela Ortems e ajudar no troubleshooting de diferenças de schema.

### Seleção de módulos Ortems

A aba **Modules** permite configurar o escopo funcional do cenário:

- **PS - Production Scheduling** ou **MP - Master Planning**;
- **SRP**;
- **WO Links**;
- **Complex Routings**;
- **Secondary Resources**;
- **Batch Machines**;
- **Inventory**;
- **Markers**;
- **Limited Resources**;
- **Parameters**.

As abas são habilitadas ou bloqueadas automaticamente de acordo com os módulos selecionados.

### Edição dos dados pela interface

Cada aba possui botões para:

- adicionar linha;
- editar linha selecionada;
- excluir linha selecionada;
- excluir todas as linhas da aba;
- importar CSV;
- exportar CSV;
- carregar exemplos.

As principais abas são:

| Aba | Finalidade | Principais tabelas Ortems envolvidas |
|---|---|---|
| Calendars | Calendários e períodos de trabalho | `B_CAL`, `B_PERI` |
| Machines | Sites, centros de trabalho, seções e máquinas | `B_ZONE`, `B_ILOT`, `B_SECT`, `B_MACH` |
| Operations | Operações e cadências por máquina | `B_OPE`, `B_CADE` |
| Routings | Gamas/roteiros e fases | `B_GAMM`, `B_PHAS` |
| Items | Itens, versões e roteiros de produção | `B_ART`, `B_VER_ART`, `E_GAMME_NOME` quando aplicável |
| BOM | Estrutura de produto | `B_NOME` |
| Work Orders | Ordens de produção planejadas/lançáveis | `B_OF` |
| WO Links | Relações de precedência entre WOs/operações | `B_PROF`, quando aplicável após lançamento |
| Secondary Resources | Recursos secundários por operação | tabelas auxiliares de recursos, conforme schema |
| Capacity | Calendários de capacidade | tabelas de capacidade conforme schema |
| Inventory Movements | Movimentos de estoque | tabelas de estoque conforme schema |

### Importação e exportação Excel

O app permite exportar um workbook estruturado com as abas de dados e importar esse mesmo workbook de volta.

Regras importantes:

- não renomeie as abas geradas;
- não altere os cabeçalhos;
- os dados são tratados como texto para evitar conversões automáticas de data/hora pelo Excel;
- a coluna `Line` é regenerada pelo app durante a importação.

O workbook exportado contém uma aba `README` com status de integridade no momento da exportação.

### Validação de integridade

Antes de gerar ou executar SQL, o Toolbox executa uma validação cruzada entre as abas. Exemplos:

- operações apontando para centros de trabalho e máquinas existentes;
- roteiros apontando para operações existentes;
- itens manufaturados apontando para roteiros válidos;
- BOM apontando para itens, versões e roteiros válidos;
- Work Orders apontando para itens e roteiros válidos;
- movimentos de estoque apontando para itens válidos;
- formatos de calendário, capacidade e datas.

Quando possível, o app oferece correções automáticas simples, como normalização de referências por diferença de maiúsculas/minúsculas, espaços ou seleção única possível.

### Geração e execução de SQL

A aba **Generate & Run SQL** permite:

- gerar o SQL completo;
- revisar o script antes da execução;
- executar no banco conectado;
- salvar o script em arquivo `.sql`;
- consultar o log de execução.

Por padrão, o SQL é executado dentro de transação. Em caso de erro, o app tenta fazer rollback para evitar banco parcialmente carregado.

---

## Comportamento atual das Work Orders

A versão atual importa Work Orders como ordens **lançáveis**, criando registros em `B_OF`.

Ela **não** pré-cria registros em `E_OF` nem em `B_BT`.

Isso é intencional.

No Ortems, `E_OF` e `B_BT` pertencem ao lado de lançamento/WIP do modelo e devem ser criados pelo fluxo nativo de **Launching** do próprio Ortems. Quando esses registros são criados manualmente durante a importação, o Ortems pode entender que a WO já existe no WIP e retornar erros como:

```text
E_OF_EXISTS
```

Por isso, o fluxo correto é:

1. importar os dados pelo Toolbox;
2. abrir o Ortems;
3. revisar as Work Orders;
4. lançar as Work Orders pelo processo nativo do Ortems.

### WO Links

Os links entre Work Orders dependem de linhas operacionais lançadas, normalmente associadas a `B_BT`/`B_PROF`, dependendo da versão/schema Ortems.

Como o Toolbox mantém as WOs em modo lançável, os WO Links são pulados durante a importação de WOs não lançadas. Se o cenário exigir links por operação, crie ou atualize esses vínculos depois que o Ortems tiver lançado as WOs e gerado as linhas correspondentes.

---

## Modo de clean import

A versão atual força um modo de importação limpa global para evitar resíduos antigos no banco de demonstração.

Durante a geração do SQL, o app:

1. desabilita constraints em todas as tabelas de usuário;
2. verifica se alguma FK continuou habilitada;
3. desabilita triggers;
4. apaga os dados de todas as tabelas de usuário não sistêmicas;
5. valida se as tabelas ficaram vazias;
6. recria os dados necessários ao cenário;
7. recria dados de referência exigidos pelas Work Orders, como série e en-cours padrão;
8. reabilita triggers;
9. revalida constraints.

> ⚠️ Esse processo é destrutivo. Ele foi desenhado para bases de demo/sandbox. Faça backup antes de executar.

---

## Requisitos

### Para executar

- Windows 10 ou superior.
- Acesso a uma base Ortems em SQL Server.
- SQL Server ODBC/OLE DB provider disponível na máquina.
- Permissões de banco suficientes para:
  - conectar;
  - inserir/atualizar/deletar dados;
  - desabilitar/reabilitar constraints e triggers, quando o clean import for usado.
- Microsoft Excel instalado, apenas se for usar importação/exportação de workbook.

### Para compilar

- AutoIt 3.
- SciTE4AutoIt3 ou AutoIt3Wrapper.
- Compilação x64 habilitada.
- Arquivos auxiliares do projeto na mesma pasta do código-fonte.

---

## Estrutura sugerida do repositório

```text
.\Toolbox
├── Toolbox.au3
├── Updater_lib2.au3
├── Toolbox.ico
├── splash.jpg
├── Help.html
├── Updater.exe
├── FileUpdate.exe
├── README.md
└── .gitignore
```

Arquivos gerados em runtime, como `settings.ini`, logs e arquivos exportados, não devem ser versionados.

---

## Como compilar

1. Clone ou copie o projeto para uma pasta local:

   ```text
   .\Toolbox
   ```

2. Abra `Toolbox.au3` no SciTE/AutoIt.

3. Garanta que os arquivos auxiliares estejam na mesma pasta:

   ```text
   .\Updater_lib2.au3
   .\Toolbox.ico
   .\splash.jpg
   .\Help.html
   .\Updater.exe
   .\FileUpdate.exe
   ```

4. Antes de publicar ou compilar para distribuição, use caminhos relativos nas diretivas do AutoIt3Wrapper. Exemplo:

   ```autoit
   #AutoIt3Wrapper_Res_File_Add=.\Updater.exe
   #AutoIt3Wrapper_Res_File_Add=.\Help.html
   #AutoIt3Wrapper_Run_After=.\FileUpdate.exe
   ```

5. Compile como x64.

6. Gere um release contendo o executável final e, se necessário para o modelo de atualização usado no projeto, os arquivos de suporte do updater.

---

## Como usar

### 1. Conectar ao banco

1. Abra o Toolbox.
2. Na aba **Database**, informe:
   - Server / Instance;
   - Database name;
   - tipo de autenticação.
3. Clique em **DB Connect**.
4. Confirme que o status mudou para conectado.

### 2. Selecionar módulos

1. Vá para a aba **Modules**.
2. Selecione PS ou MP.
3. Marque os módulos necessários para o cenário.
4. Clique em **Apply / Show Summary**, se quiser revisar o escopo.

### 3. Preencher dados

Preencha as abas na sequência lógica:

1. Calendars;
2. Machines;
3. Operations;
4. Routings;
5. Items;
6. BOM;
7. Work Orders;
8. módulos opcionais, quando aplicável.

Você também pode usar **Load Example** para criar uma base inicial de exemplo e depois ajustar os dados.

### 4. Validar integridade

Clique em **Integrity Check**.

Corrija os problemas indicados antes de gerar o SQL. Quando o app oferecer auto-fix, revise o resultado antes de continuar.

### 5. Gerar SQL

Clique em **Generate SQL**.

Revise o conteúdo na aba **Generate & Run SQL**. Essa etapa é importante porque o script pode deletar e recriar dados da base conectada.

### 6. Executar no banco

Clique em **Run on DB**.

Confirme a execução apenas se o banco conectado for o banco correto e se houver backup ou possibilidade de recriação do ambiente.

### 7. Abrir no Ortems

Depois da execução bem-sucedida:

1. abra o Ortems;
2. verifique os dados importados;
3. lance as Work Orders pelo fluxo nativo do Ortems.

---

## Arquivo `settings.ini`

O app cria automaticamente um arquivo `settings.ini` na mesma pasta do executável/código.

Exemplo:

```ini
[Connection]
Server=localhost\SQLEXPRESS
Database=ORTEMS_DEMO
Auth=Windows Authentication
User=

[Modules]
Mode=PS
SRP=0
WOL=0
SR=0
INV=0
MRK=0
LR=0
PRM=0
BATCH=0
CROUT=0

[SQL]
ClearFirst=1
Transaction=1
Verbose=0

[Logging]
VerboseMode=0

[Window]
X=-1
Y=-1
W=1100
H=720
```

Notas:

- o arquivo é local por máquina/usuário;
- não publique esse arquivo no GitHub;
- a senha não deve ser versionada;
- para troubleshooting, ative `VerboseMode=1` ou marque **Verbose Mode** na interface.

---

## Logs e troubleshooting

Os logs são gravados em:

```text
.\log
```

Use **Verbose Mode** quando precisar investigar:

- erro de conexão;
- coluna não encontrada em alguma tabela Ortems;
- erro de FK;
- rollback de transação;
- problema na geração de SQL;
- erro durante clean import.

O modo verbose registra detalhes adicionais, incluindo consultas de detecção de schema e comandos executados.

### Erro `E_OF_EXISTS` ao lançar WOs

Esse erro normalmente indica que a base já contém registros de lançamento/WIP relacionados à mesma WO.

A versão atual evita esse problema por design:

- cria a WO em `B_OF`;
- não cria `E_OF`;
- não cria `B_BT`;
- limpa resíduos antigos durante o clean import.

Se o erro continuar acontecendo, confirme que:

1. você está usando a versão atual do app;
2. o SQL foi gerado novamente depois da atualização;
3. a importação foi executada com sucesso;
4. não há customização externa recriando `E_OF` ou `B_BT` antes do lançamento no Ortems;
5. o log não indica falha no clean import ou rollback.

### Erro de FK após importação

Verifique:

- se todas as abas passaram no **Integrity Check**;
- se os IDs estão escritos exatamente como esperado;
- se a versão/schema do Ortems possui nomes de colunas diferentes;
- se o log verbose mostra alguma coluna não detectada;
- se o banco tinha triggers/customizações específicas.

### Excel não abre na importação/exportação

A importação/exportação de workbook usa automação COM do Microsoft Excel. O Excel precisa estar instalado na máquina Windows que executa o app.

---

## Boas práticas

- Use sempre uma base clonada ou de demonstração.
- Faça backup antes de executar **Run on DB**.
- Gere e revise o SQL antes de rodar.
- Mantenha nomes de IDs simples, estáveis e sem caracteres especiais desnecessários.
- Preencha os dados na ordem lógica: calendário → máquina → operação → roteiro → item → BOM → WO.
- Use **Integrity Check** antes de qualquer execução.
- Ative **Verbose Mode** antes de reproduzir um erro.
- Não versionar `settings.ini`, logs, workbooks de cliente ou exports com dados sensíveis.

---

## `.gitignore` sugerido

```gitignore
# Runtime settings
settings.ini

# Logs
log/
*.log

# Build output
*.exe
*.a3x
*.bak

# Local exports/imports
*.xlsx
*.xlsm
*.csv
*.sql

# Editor/OS files
.vscode/
.idea/
.DS_Store
Thumbs.db
```

> Se você quiser publicar binários pelo GitHub Releases, mantenha os `.exe` fora do repositório principal e anexe-os ao release.

---

## Segurança e privacidade

Antes de publicar este projeto:

- remova caminhos absolutos da sua máquina;
- use apenas caminhos relativos, como `.\Toolbox`;
- não publique `settings.ini`;
- não publique logs reais de clientes;
- não publique exports Excel/CSV/SQL contendo dados sensíveis;
- revise metadados do executável antes de distribuir;
- valide se o updater não contém tokens, URLs privadas ou credenciais.

---

## Limitações conhecidas

- Aplicação Windows-only.
- Requer SQL Server e provider ODBC/OLE DB disponível.
- Import/export Excel requer Microsoft Excel instalado.
- O schema Ortems pode variar entre versões; o app tenta detectar colunas dinamicamente, mas schemas muito customizados podem exigir ajuste no código.
- WO Links por operação não são aplicados durante importação de WOs lançáveis; eles devem ser tratados depois que o Ortems criar as linhas de lançamento.
- O clean import é destrutivo e não deve ser usado em produção.

---

## Roadmap sugerido

Ideias para evolução futura:

- opção de modo não destrutivo para refresh parcial;
- tela dedicada para diagnóstico de schema Ortems;
- relatório HTML de integridade;
- exportação/importação JSON além de Excel/CSV;
- suporte mais completo a WO Links pós-lançamento;
- tela de preview do impacto no banco antes do `Run on DB`;
- testes automatizados para geração SQL por módulo.

---

## Licença

Defina a licença antes de publicar o projeto.

Sugestão: incluir um arquivo `LICENSE` na raiz do repositório e atualizar esta seção com o nome da licença escolhida.

---

## Aviso

Este projeto manipula diretamente tabelas de banco Ortems SQL Server. Use com cuidado, sempre em ambiente controlado, e revise o SQL gerado antes da execução.
