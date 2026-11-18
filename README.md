# Tracking Certificate

Aplicativo desktop (Python + PySide6/Qt) da equipe de Qualidade de Fornecedores
(Bosch Rexroth, planta PUQ) que **monitora a validade dos certificados dos
fornecedores brasileiros** (ISO 9001, IATF 16949 e ISO 14001/Ambiental) a partir
da exportação do SRM em Excel e **envia e-mails de cobrança pelo Outlook**, no
manual ou de forma 100% automática.

> Interface em inglês. Conteúdo dos e-mails, documentação e comentários do código
> em português.

---

## Sumário

**Parte 1 — Manual do usuário** (como usar o aplicativo)
1. [Visão geral](#1-visão-geral)
2. [Primeiros passos](#2-primeiros-passos)
3. [Configuração inicial (Settings)](#3-configuração-inicial-settings)
4. [Fluxo de trabalho diário](#4-fluxo-de-trabalho-diário)
5. [Telas do aplicativo](#5-telas-do-aplicativo)
6. [Regras de negócio](#6-regras-de-negócio)
7. [E-mails enviados](#7-e-mails-enviados)
8. [Execução automática (Agendador de Tarefas)](#8-execução-automática-agendador-de-tarefas)
9. [Arquivos gerados](#9-arquivos-gerados)
10. [Perguntas frequentes e solução de problemas](#10-perguntas-frequentes-e-solução-de-problemas)

**Parte 2 — Manual do desenvolvedor** (instalação, código, arquitetura)
11. [Obter o código e preparar o ambiente](#11-obter-o-código-e-preparar-o-ambiente)
12. [Arquitetura](#12-arquitetura)
13. [Estrutura do projeto](#13-estrutura-do-projeto)
14. [Módulos em detalhe](#14-módulos-em-detalhe)
15. [Modelo de dados](#15-modelo-de-dados)
16. [Fluxo de execução](#16-fluxo-de-execução)
17. [Interface (Qt)](#17-interface-qt)
18. [Como estender o app](#18-como-estender-o-app)
19. [Testes e depuração](#19-testes-e-depuração)
20. [Decisões de projeto e limitações conhecidas](#20-decisões-de-projeto-e-limitações-conhecidas)
21. [Histórico de mudanças](#21-histórico-de-mudanças)

---

# PARTE 1 — MANUAL DO USUÁRIO

## 1. Visão geral

O app faz quatro coisas:

1. **Lê** o Excel exportado do SRM (uma linha por fornecedor, com as datas de
   validade dos certificados).
2. **Classifica** cada fornecedor em cada certificado: Ativo, Expirando
   (90/60/30 dias), Vencido ou Sem data.
3. **Cobra** os fornecedores por e-mail (Outlook), sem repetir avisos
   desnecessários.
4. **Acompanha** o que já foi tratado: histórico de envios, certificados
   recebidos, fornecedores em Risk Assessment e relatórios gráficos.

Só são processados fornecedores com **Country Code = BR**; os demais são
descartados na leitura do Excel.

## 2. Primeiros passos

Com o aplicativo já instalado (a equipe de TI/desenvolvimento cuida da instalação — ver
[seção 11](#11-obter-o-código-e-preparar-o-ambiente)), abra o **Tracking Certificate** e:


1. Vá em **Settings** e preencha o que for necessário (seção 3).
2. Clique em **Save Settings**.
3. No **Panel**, clique em **1. Select SRM Excel File** e escolha o arquivo.
4. Clique em **2. Load & Analyze**.

Nas próximas aberturas, o último Excel escolhido já vem selecionado.

## 3. Configuração inicial (Settings)

A tela **Settings** (rodapé da barra lateral) tem os grupos abaixo. Todos os
valores são salvos em `tracking_config.json` pelo botão **Save Settings**.

### 3.1 SRM Excel Columns

Nomes exatos das colunas do Excel do SRM. Se o SRM mudar um nome, ajuste aqui.

| Campo | Padrão |
|---|---|
| Sheet name (opcional) | vazio = procura em todas as abas |
| Supplier code (RBSNO) | `Basic Data: RBSNO` |
| Supplier code (PMD/SAP) | `Athena Master` |
| Supplier name | `Supplier Name` |
| ISO 9001 expiry | `Certificates: Quality ISO 9001 and others: Expiry date` |
| IATF 16949 expiry | `Certificates: Quality IATF 16949: Expiry date` |
| Environmental expiry | `Certificates: Environment: Expiry date` |
| Country code | `Country Code` (só para filtrar BR) |

O leitor localiza sozinho a linha de cabeçalho (varre as primeiras 40 linhas) e
aceita cabeçalhos com variações de caixa/espaços. Todas as colunas acima
precisam existir; se alguma faltar, o erro lista o que foi encontrado.

### 3.2 Supplier E-mail Directory (Access)

De onde vêm os e-mails dos fornecedores.

| Campo | Padrão |
|---|---|
| Access database (.accdb) | (escolher o arquivo) |
| Table | `tbl_Fornecedores` |
| RBSNO column | `RBSNO` |
| PMD column | `PMD Master` |
| E-mail column | `Emails Concatenados` |

O e-mail é localizado primeiro pelo **RBSNO** e, se não achar, pelo **PMD**.
O campo de e-mail pode conter vários endereços (concatenados) — o valor é
usado como está no campo *Para* do Outlook.

### 3.3 Certificate Files

**Folder for certificate files:** pasta-base onde ficam os certificados
recebidos e os arquivos do Risk Assessment. O app cria automaticamente uma
subpasta por fornecedor no formato `Nome do fornecedor - RBSNO`. As evidências
do Risk Assessment ficam na subpasta `Evidence` dentro dela.

**Nomes dos certificados:** todo certificado anexado (pelo Register ou pelo Risk
Assessment) é renomeado para `<Certificado> - <Fornecedor> - <RBSNO>.<ext>`
(ex.: `IATF 16949 - ZFR COM PERFILADOS LTDA - 97942806.pdf`; prefixos `ISO 9001`,
`IATF 16949` e `ISO 14001`). **O app reconhece o certificado pelo nome do
arquivo**, então não renomeie nem mova esses arquivos.

### 3.3.1 Shared Database (banco compartilhado)

**Database file (tracking.db):** local do arquivo do banco (histórico de envios,
Risk Assessment e certificados recebidos). Use o botão **...** para selecionar um
`.db` existente ou escolher onde ele será criado (se o arquivo não existir, é
criado vazio no local indicado).

**Para vários PCs usarem os mesmos dados:** em **cada PC**, aponte esta opção
para o **mesmo arquivo** em uma pasta de rede (ex.: `\\servidor\qualidade\tracking.db`).
A configuração é gravada por PC (`tracking_config.json`), então é preciso
informá-la uma vez em cada máquina. Para migrar dados já existentes, copie o
`tracking.db` atual para o local compartilhado **antes** de apontar os PCs para
ele. Ao salvar, o app valida a pasta (precisa existir e ser acessível) e cria as
tabelas se o banco for novo. Evite que dois PCs enviem os lembretes ao mesmo
tempo (o histórico é gravado a cada envio, mas dois envios simultâneos podem
duplicar e-mails).

### 3.4 Notification Rules

| Campo | Padrão | Significado |
|---|---|---|
| Reminder 1 / 2 / 3 (dias) | 90 / 60 / 30 | Faixas de aviso antes do vencimento |
| Resend expired certificate notice every (days) | **15** | Intervalo de reenvio para vencidos e para certificados em branco |

### 3.5 E-mail Content

| Campo | Padrão | Significado |
|---|---|---|
| Company name | `Bosch Rexroth LTDA.` | Aparece no cabeçalho do e-mail |
| BCC address | `Digital.Solutions@boschrexroth.com.br` | **Cópia oculta**: o fornecedor não vê |
| Default subject (renewal / expired) | `Renovação de Certificado - Ação Necessária` | Prefixo do assunto de renovação |
| Subject when ISO 9001 and IATF 16949 are blank | `Confirmação de Certificado - Ação Necessária` | Prefixo do assunto de "certificado em branco" |
| Sender address | vazio | Remetente (ver abaixo) |
| **Preview E-mail Content...** | — | Abre a pré-visualização de cada modelo de e-mail |

**Sender address (remetente).** Em branco, o e-mail sai pela conta padrão do
Outlook do usuário. Para enviar por outra caixa (ex.:
`Digital.Solutions@boschrexroth.com.br`):

- se o endereço for uma **conta configurada neste Outlook**, o app envia por ela;
- caso contrário, usa "enviar em nome de" (`SentOnBehalfOfName`), que exige que
  o administrador do Exchange conceda ao seu usuário a permissão **Send As** ou
  **Send on Behalf** naquela caixa. Sem a permissão o envio falha e o motivo
  aparece no **Log**.

## 4. Fluxo de trabalho diário

1. Baixe o Excel atualizado do SRM.
2. **Panel → Select SRM Excel File** (se o arquivo mudou) → **Load & Analyze**.
3. Confira as abas **ISO 9001 / Environmental / IATF 16949** e o contador
   "N pending".
4. (Opcional) **Preview E-mail Content...** em Settings para revisar os textos.
5. Clique em **Send Pending Reminders** e confirme.
6. Acompanhe em **Send History** e **Log**.
7. Fornecedor mandou o arquivo mas ainda não subiu no datalake? Use
   **Register Received Certificate** (seção 5.2).
8. Fornecedor sem o certificado? Cadastre em **Risk Assessment** (seção 5.3).

## 5. Telas do aplicativo

A **barra lateral** pode ser recolhida/expandida pelo botão de menu (☰). Recolhida,
mostra só os ícones. Ordem: **Panel, Risk Assessment, Reports** no topo e
**Send History, Log, Settings** na base.

### 5.1 Panel

- **Select SRM Excel File / Load & Analyze:** escolhe e processa o Excel (o último
  arquivo escolhido é lembrado).
- **Register Received Certificate:** ver 5.2.
- **4 cartões** (Suppliers, Active, Expiring, Expired) que mostram os números
  **da sub-aba selecionada**.
- **Campo de pesquisa** (canto direito, acima da tabela): filtra as 3 abas por
  **Supplier, RBSNO ou PMD** enquanto você digita.
- **Sub-abas** ISO 9001 (Quality), Environmental (ISO 14001), IATF 16949 (Quality).
  Cada tabela tem as colunas: Supplier, RBSNO, PMD, Expiry Date, Status,
  Reminder, E-mail. Clique no cabeçalho para ordenar (clicar de novo inverte).
  A ordenação de **Status** é alfabética por tipo (Active, Expired, Expires,
  Missing date, Risk Assessment, Supplier IATF/ISO).
- **Send Pending Reminders:** envia os e-mails pendentes (pede confirmação).
  O contador ao lado mostra quantos estão pendentes.

**Cores das linhas**

| Cor | Situação |
|---|---|
| Verde | Ativo (mais de 90 dias) ou certificado recebido manualmente |
| Amarelo | Expirando (faixa 90/60/30) |
| Vermelho | Vencido |
| Cinza | Missing date (sem data) |
| Azul | Em dia pelo **outro** certificado (Supplier ISO / Supplier IATF) |
| Roxo | Fornecedor em **Risk Assessment** |

**Coluna Reminder:** `Pending` (será enviado), `Sent` (já enviado nesta faixa),
`Received` (certificado recebido manualmente, conta como Ativo), `-` (nada a fazer).

### 5.2 Register Received Certificate

Para quando o fornecedor **já enviou o arquivo**, mas ele ainda não foi
carregado no datalake — evita cobrar em duplicidade.

1. Clique no botão no Panel (exige ter analisado o Excel antes).
2. Escolha o fornecedor no dropdown (digite para filtrar por qualquer parte do
   nome). O app mostra o RBSNO e o status/data dos 3 certificados.
3. Informe **qual certificado** está anexando (ISO 9001, IATF 16949 ou
   Environmental).
4. Anexe o **arquivo (obrigatório)**.
5. **Save:** o arquivo é copiado para `<pasta>/<Fornecedor> - <RBSNO>/` (criada
   se não existir) e **renomeado** para `<Certificado> - <Fornecedor> - <RBSNO>.<ext>`
   (ex.: `ISO 9001 - Hidro Jet - 902193213.pdf`, independentemente do nome
   original). Registrar de novo o mesmo certificado do mesmo fornecedor
   **substitui** o arquivo e o registro anteriores. O app reanalisa em seguida.

Efeito: o certificado passa a contar como **Ativo** (linha verde, Status "Active",
Reminder `Received`) e deixa de ser cobrado. Registrar o **ISO 9001** faz a aba
IATF 16949 ficar azul ("Supplier ISO"), e registrar o **IATF 16949** faz a aba
ISO 9001 ficar azul ("Supplier IATF"). Registrar o Environmental afeta só o
ambiental.

**Como o app confere:** a cada **Load & Analyze** (e no `--auto`) ele
1. apaga os registros cujo arquivo **não existe mais** na pasta (arquivo
   apagado = fornecedor volta ao status real);
2. também reconhece arquivos colocados na pasta do fornecedor com os nomes
   acima, mesmo sem passar pelo botão;
3. encerra o registro de um certificado quando o Excel passa a mostrá-lo
   **válido** (a data da planilha volta a valer).

O app **não lê a data de validade do arquivo**: o certificado fica Ativo
enquanto o arquivo existir (apague-o quando o certificado vencer).

> Requer a pasta configurada em **Settings → Certificate Files**.

### 5.3 Risk Assessment

Lista de fornecedores que **não possuem o certificado** e exigem ação em campo.

- **Add / Edit / Remove** (Edit também por duplo clique).
- No cadastro: nome (dropdown com os fornecedores analisados — preenche RBSNO
  e PMD), certificado faltante, **status** (Open / In progress / Completed),
  observações e dois campos de arquivos:
  - **Certificate:** o certificado do fornecedor. É renomeado no padrão
    `<Certificado> - <Fornecedor> - <RBSNO>.<ext>` e funciona como o Register
    (o fornecedor passa a Ativo no próximo Load & Analyze). Para "ISO 9001 or
    IATF 16949" o nome usa o prefixo ISO 9001 — se o arquivo for o IATF, escolha
    "IATF 16949" no campo antes de salvar.
  - **Evidence:** evidências do assessment — **arquivos** (copiados para a
    subpasta `Evidence`, com o nome original) e/ou **links** (hiperlinks; duplo
    clique abre o arquivo ou o link).
- **Remove selected** (em Certificate ou Evidence) tira o item da lista; o
  arquivo só é **apagado da pasta ao clicar em Save** (Cancel não apaga nada).
- **Histórico:** abrir um novo Risk Assessment para o mesmo fornecedor e
  certificado cria **outro registro**; os anteriores (inclusive Completed) são
  mantidos. Arquivos de evidência com nome repetido recebem data/hora, sem
  sobrescrever.
- **Busca** por nome ou RBSNO e **filtros** por certificado e status; todas as
  colunas ordenam ao clicar no cabeçalho.
- **Open folder** abre a pasta do fornecedor.

Efeito no Panel: enquanto houver Risk Assessment **em aberto** (Open ou In
progress), o fornecedor fica **roxo**, com status **Risk Assessment**, e **não
recebe e-mails de cobrança** de ISO 9001/IATF 16949. Um Risk Assessment
**Completed** fica só no histórico: não pausa a cobrança e não entra na
contagem do Follow-up.

### 5.4 Reports

- **Rosca "Quality"** e **rosca "Environmental"**: distribuição Ativo /
  Expirando / Vencido / Sem data. (IATF não tem rosca própria — completa o ISO
  9001, por isso "Quality".)
- **Comparison by certificate:** barras de Quality x Environmental.
- **Follow-up:** quantos fornecedores estão em **Risk Assessment**, quantos
  tiveram **arquivo adicionado manualmente** e quantos estão em
  **Missing date (Quality)**. *Missing date = total sem data (sem os já
  recebidos) − Risk Assessment em aberto.*
- **Time Saved** (selo verde, canto superior direito): cada e-mail enviado = 5
  minutos economizados, exibido em horas.

### 5.5 Send History

Todos os e-mails já enviados (RBSNO, PMD, fornecedor, certificado, validade,
faixa, data, destinatário), com botão **Refresh**, o contador **total de
envios** (à direita, rente à tabela). O **Time Saved** fica na aba Reports
(cada envio = 5 minutos; ex.: 20 envios = 1,67 h).

### 5.6 Log

Registro com data e hora de tudo que acontece: arquivo selecionado, início e
fim de análises, erros (Excel, Access, Outlook), envios feitos/falhos/pulados,
configurações salvas, cadastros de Risk Assessment e certificados recebidos,
resultado das notificações do Windows. É gravado em `tracking_app_log.txt` e as
últimas 500 linhas são recarregadas ao abrir o app.

## 6. Regras de negócio

### 6.1 Classificação de um certificado

| Situação | Condição |
|---|---|
| Missing date | célula de data em branco |
| Vencido | data anterior a hoje |
| Expirando | faltam ≤ 90, ≤ 60 ou ≤ 30 dias (faixa mais próxima) |
| Ativo | mais de 90 dias |

### 6.2 Quando um e-mail é enviado

- **Faixas 90/60/30:** um e-mail **por faixa** (nunca repete a mesma faixa).
- **Vencido:** e-mail ao entrar em vencido e depois **a cada N dias** (padrão 15)
  até o Excel mostrar uma nova data de validade futura.
- **ISO 9001 e IATF 16949 ambos em branco:** e-mail de **confirmação** (ver 7.2),
  reenviado a cada N dias até o fornecedor ser tratado.
- **Ambiental em branco: nunca gera e-mail.**
- Fornecedor **sem e-mail** no Access: é pulado e registrado no Log.
- O controle é feito pelo histórico em `tracking.db`.

### 6.3 ISO 9001 × IATF 16949 (substituição cruzada)

Basta ter **um dos dois** válido. Ambiental é totalmente independente.

| Situação do fornecedor | Aba ISO 9001 | Aba IATF 16949 |
|---|---|---|
| ISO ativo, IATF qualquer | normal | **azul** "Supplier ISO" (conta como Ativo) |
| IATF ativo, ISO qualquer | **azul** "Supplier IATF" (Ativo) | normal |
| ISO expirando, IATF ativo | azul "Supplier IATF" (Ativo, sem aviso) | normal (Ativo) |
| ISO válido (expirando) e IATF vencido/em branco | normal (expirando) | azul "Supplier ISO" |
| Ambos expirando | normal (expirando) | normal (expirando) |
| **Nenhum válido** | vermelho/cinza + e-mail | vermelho/cinza (sem e-mail próprio) |

Quando **nenhum dos dois é válido**, é enviado **um único e-mail** pedindo
"ISO 9001 ou IATF 16949" (o envio fica atrelado à linha da aba ISO 9001; a linha
da aba IATF só espelha o status), para nunca duplicar a cobrança.

> Consequência: se um certificado está expirando mas o outro está ativo, **não
> sai o aviso de 90/60/30 dias** daquele certificado — o outro "repõe" o primeiro.

### 6.4 Prioridade entre estados especiais

Certificado **recebido** (arquivo na pasta) tem prioridade sobre `Risk Assessment`:
a linha vira Ativa/`Received`. Risk Assessment **em aberto** pausa a cobrança de
ISO 9001/IATF 16949 (não afeta o Ambiental); concluído, não pausa.

## 7. E-mails enviados

Sempre em **português**, HTML com identidade azul, assinatura "Equipe de
Qualidade de Fornecedores". Assunto:
`<prefixo> - <tag do certificado> - <RBSNO do fornecedor>` (ex.: `Renovação de
Certificado - Ação Necessária - ISO 9001 / IATF 16949 - 9092345`; o RBSNO no fim
permite automações futuras), com a tag em
`ISO 9001`, `ISO 14001`, `IATF 16949` ou `ISO 9001 / IATF 16949`.

### 7.1 Renovação (90/60/30 e vencido)

Informa certificado, data de validade, dias restantes (ou dias de atraso) e
orienta o fornecedor a enviar o certificado renovado. Tom e destaque variam por
faixa (aviso x urgente x vencido).

### 7.2 Certificado em branco (ISO 9001 e IATF 16949)

Pergunta ao fornecedor se ele possui o certificado:

- **se possui:** responder o e-mail anexando o arquivo;
- **se não possui:** informar na resposta, para seguirmos com o **Risk Assessment**.

### 7.3 Pré-visualização

**Settings → Preview E-mail Content...** mostra assunto e corpo para cada
combinação de certificado (ISO 9001, Ambiental, IATF 16949, "ISO 9001 ou IATF
16949") e faixa (90, 60, 30, vencido, certificado em branco).

## 8. Execução automática (Agendador de Tarefas)

O envio normal é manual. Para rodar sozinho:

```
python main.py --auto
```

Sem interface e sem confirmação: lê as configurações salvas, analisa o Excel
(último arquivo escolhido), respeita Received/Risk Assessment, envia os
pendentes e grava tudo em `tracking_auto_log.txt`. Ao final dispara uma
notificação do Windows com o resumo.

**Como agendar**

1. Abra o app uma vez, configure tudo em Settings, escolha o Excel e salve.
2. Agendador de Tarefas → **Criar Tarefa Básica** → gatilho (ex.: diário 08:00).
3. Ação **Iniciar um programa**:
   - **Programa/script:** caminho do `python.exe` (ou `pythonw.exe`), entre aspas
     se tiver espaços — ex.: `"C:\Users\voce\venv\Scripts\python.exe"`.
   - **Adicione argumentos:** `main.py --auto` (**sem aspas**, nem simples).
   - **Iniciar em:** pasta do app (**sem aspas**), ex.: `C:\TrackingCertificate`.
4. Se necessário, marque "Executar estando o usuário conectado ou não".

Requisitos: PC ligado (sem suspensão), Outlook instalado/configurado, e o
**Excel precisa estar atualizado no mesmo caminho** (o modo automático lê o
arquivo do caminho salvo; substitua o arquivo, mantendo nome e pasta).

Erros comuns: *Launch Failure / 2147942667* ou *OpenWith.exe* indicam
caminho do `main.py` colocado no campo "Programa/script", aspas simples ou
caractere sobrando no "Iniciar em".

## 9. Arquivos gerados

Todos ao lado do app (mesma pasta de `main.py`):

| Arquivo | Conteúdo |
|---|---|
| `tracking_config.json` | Configurações (e último Excel) |
| `tracking.db` | SQLite: histórico de envios, Risk Assessment, certificados recebidos (local configurável em Settings → Shared Database) |
| `tracking_app_log.txt` | Log da interface |
| `tracking_auto_log.txt` | Log do modo `--auto` |
| `<pasta de certificados>/<Fornecedor> - <RBSNO>/` | Certificados (`ISO 9001 - …`, `IATF 16949 - …`, `ISO 14001 - …`) e subpasta `Evidence` |

Arquivos antigos `renovacert_*` são renomeados automaticamente na primeira
execução (config, banco, logs), preservando o histórico.

**Backup:** copie `tracking.db` e `tracking_config.json`.

## 10. Perguntas frequentes e solução de problemas

**O Excel dá erro de coluna não encontrada.** Compare os nomes em
*Settings → SRM Excel Columns* com o cabeçalho do arquivo. Se a exportação não
trouxer a coluna do IATF, o leitor acusa a ausência — ajuste o nome ou inclua a
coluna no relatório do SRM.

**Erro "too few parameters" ao ler o Access.** O nome de uma coluna em
*Supplier E-mail Directory* não existe na tabela. A mensagem de erro lista as
colunas reais — copie o nome exato.

**Fornecedor aparece "(no e-mail on file)".** RBSNO/PMD não encontrados no
Access, ou colunas configuradas erradas. Sem e-mail não há envio.

**O e-mail saiu pela conta errada / falhou com "Sender".** Veja 3.5: a conta
precisa existir no Outlook ou o usuário precisa de permissão Send As/Send on
Behalf. O Log mostra o erro do Outlook.

**"Outlook automation is not available".** O pywin32 não está instalado ou o
app não está rodando no Windows com Outlook.

**Notificações do Windows não aparecem.** Instale `plyer` (`pip install
plyer`). Se falhar, o Log registra `Desktop notification NOT shown (...)` com o
motivo; o app continua funcionando normalmente.

**Reenviou aviso que eu achava já ter enviado.** O histórico é por
RBSNO + tipo de certificado + data de validade. Se a data de validade mudou no
Excel, conta como novo ciclo. Se o `tracking.db` foi apagado/trocado, o
histórico recomeça.

**Quero parar de cobrar um fornecedor.** Arquivo recebido → *Register Received
Certificate* (ou anexe em *Certificate* no Risk Assessment). Sem certificado →
*Risk Assessment* (em aberto).

**Anexei o certificado mas o fornecedor continua vermelho/cinza.** O app
reconhece pelo **nome** do arquivo (`ISO 9001 - …`, `IATF 16949 - …`,
`ISO 14001 - …`) dentro da pasta do fornecedor, configurada em Settings →
Certificate Files. Clique em **Load & Analyze** de novo. Arquivos antigos com
outro nome são renomeados ao abrir o registro no Risk Assessment e clicar em Save.

**Apaguei o arquivo e o fornecedor continua verde.** Clique em **Load & Analyze**:
registros sem arquivo na pasta são removidos na análise.

---

# PARTE 2 — MANUAL DO DESENVOLVEDOR

## 11. Obter o código e preparar o ambiente

### Obter o código

O projeto é distribuído como pasta/zip (`RenovaCert.zip`). Extraia em uma pasta
sem restrições de escrita (o app grava config, banco e logs ao lado de
`main.py`), por exemplo `C:\TrackingCertificate`.

### Requisitos

| Item | Detalhe |
|---|---|
| Sistema | Windows (o envio de e-mail usa o Outlook via COM) |
| Python | 3.10 ou superior |
| Outlook | Instalado e com conta configurada (para enviar) |
| Driver Access | *Microsoft Access Database Engine* (ODBC) compatível com a arquitetura do Python (64 bits), para ler o `.accdb` de e-mails |

Dependências Python (`requirements.txt`):

```
PySide6==6.7.2
openpyxl>=3.1
pyodbc>=5.0
pywin32>=306
plyer>=2.1.1
```

### Instalação e execução

```
cd C:\TrackingCertificate   (ou a pasta onde extraiu)
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

> Em outros sistemas (Linux/macOS) a interface abre, mas **o envio de e-mail e a
> leitura do Access não funcionam** (dependem de Windows/Outlook/ODBC Access).

Para rodar sem interface (Agendador de Tarefas): `python main.py --auto` (ver seção 8).

### Problemas de ambiente

**`ImportError: DLL load failed while importing QtWidgets`.**
1. Se estiver no Anaconda `(base)`, use um venv limpo (`python -m venv venv`).
2. Reinstale: `pip uninstall PySide6 PySide6-Addons PySide6-Essentials shiboken6 -y`
   e depois `pip install --no-cache-dir PySide6==6.7.2`.
3. Instale o *Microsoft Visual C++ Redistributable (x64)*.
4. Se persistir, pode ser bloqueio de antivírus/EDR — acione o TI.

## 12. Arquitetura

Separação em duas camadas:

- **Lógica de negócio (raiz do projeto):** leitura de Excel/Access, regras de
  notificação, templates de e-mail, envio, bancos SQLite. **Sem dependência de
  Qt**, reutilizada pela interface e pelo modo `--auto`.
- **Interface (`ui/`):** PySide6/Qt. Só orquestra e exibe.

```
Excel SRM ─► excel_reader ─► SupplierRecord[]
Access   ─► contacts_db  ─► e-mails
tracking.db ─► history_db / received_db / risk_db
                     │
                     ▼
            supplier_rows.build_rows()   ◄── regra única (UI e --auto)
                     │  linhas por aba (iso / env / iatf)
         ┌───────────┴───────────┐
         ▼                       ▼
   ui/main_window.py        auto_run.py
   (tabelas, envio)         (envio headless)
         │                       │
         └──► templates ─► email_sender (Outlook COM) ─► history_db.log_send
```

**Princípio-chave:** toda decisão "quem recebe e-mail / que cor / que status"
vive em `supplier_rows.build_rows`. A UI e o `--auto` apenas consomem o
resultado, então os dois **nunca divergem**.

Concorrência: análise e envio rodam em `threading.Thread`; a comunicação com a
UI usa **signals do Qt** (`Worker(QObject)`), que são enfileirados para a
thread da interface. Nunca altere widgets direto de uma thread.

## 13. Estrutura do projeto

```
main.py                 Ponto de entrada (GUI ou --auto)
auto_run.py             Execução headless (Agendador de Tarefas)
config.py               DEFAULT_CONFIG, load/save, migração de arquivos antigos
excel_reader.py         Leitura do Excel do SRM (SupplierRecord)
contacts_db.py          Leitura do Access (e-mails) via pyodbc
reminder_engine.py      Faixas 90/60/30, vencido, needs_notification
supplier_rows.py        Regra central: linhas por aba, ISO x IATF, recebido, risco
templates.py            Corpos HTML dos e-mails (pt-br)
email_sender.py         Envio via Outlook COM (BCC, remetente)
history_db.py           Histórico de envios (SQLite)
received_db.py          Certificados recebidos ainda não subidos (SQLite)
risk_db.py              Risk Assessment com histórico, certificados e evidências (SQLite)
cert_files.py           Pastas por fornecedor, renomeação padrão e leitura dos arquivos
notifications.py        Notificações do Windows (Plyer, opcional)
requirements.txt
ui/
  main_window.py        Janela principal, sidebar, Panel, Settings, History, Log
  pages.py              Risk Assessment, Reports (gráficos), diálogo de recebido
  common.py             SortItem (ordenação por chave) e numeric_or_text_key
  theme.py              Cores e folha de estilos QSS
```

## 14. Módulos em detalhe

### `config.py`
`DEFAULT_CONFIG` (dict) com todas as chaves; `load_config()` mescla o JSON salvo
sobre os padrões (chaves novas aparecem sem quebrar configs antigas);
`save_config()`. `_migrate_legacy_files()` renomeia `renovacert_*` → `tracking_*`
na importação, e `load_config` corrige `history_db_path` antigo.

Chaves principais: `excel_path`, `excel_sheet_name`, `excel_col_*`
(`rbsno, pmd, name, iso, env, iatf, country`), `access_*`, `history_db_path`,
`reminder_days`, `overdue_resend_interval_days`, `company_name`, `bcc_email`,
`default_subject_pt`, `default_subject_missing_pt`, `sender_email`,
`certificates_folder`.

### `excel_reader.py`
- `SupplierRecord` (`__slots__`): `rbsno, pmd, name, iso_expiry, env_expiry,
  iatf_expiry, country_code`.
- `_find_header_row()` varre as primeiras 40 linhas e guarda a **melhor**
  correspondência (não zera a cada linha — bug antigo corrigido).
- `read_suppliers(path, columns, sheet_name, country_filter)`: busca em todas as
  abas (ou na informada), extrai e filtra por país. Lança `ValueError` com
  mensagem útil (colunas encontradas/ausentes, abas disponíveis).
- `_parse_date` aceita `datetime`, `date` e texto (`dd/mm/aaaa`, `dd-mm-aaaa`, `dd/mm/aa`, `aaaa-mm-dd`).

### `contacts_db.py`
`load_emails()` lê a tabela do Access e devolve dict com chaves
`("rbsno", x)` e `("pmd", x)`. Em falha, `_list_table_columns()` anexa as
colunas reais ao erro. `resolve_email(emails, rbsno, pmd)` tenta RBSNO e depois PMD.

### `reminder_engine.py`
- Constantes: `WARNING_THRESHOLDS`, `EXPIRED`, `NO_DATE`.
- `compute_status(expiry, today)` → `(tier, days_left)`. `tier` é `None` (ativo),
  `"90"|"60"|"30"` ou `EXPIRED`. (`NO_DATE` é atribuído por quem chama, quando a
  data é `None`.)
- `needs_notification(hdb, rbsno, cert_type, expiry, tier, interval)`:
  faixas 90/60/30 → uma vez por faixa; `EXPIRED`/`NO_DATE` → primeira vez, ou
  mudança de faixa, ou `>= interval` dias desde o último envio.
- `tier_label(tier)` → "Expired", "Missing", "N days".
- `WARNING_THRESHOLDS` é reatribuído em tempo de execução a partir da config
  (GUI e `--auto`).

### `supplier_rows.py` (núcleo)
- Constantes: `CERT_TYPES` (abas), `CERT_TYPE_LABELS_PT`, `CERT_ISO_TAG`,
  `CERT_LABEL_EN`, `COMBINED_FIELD = "combined_cert"`, `BLUE`.
- `quality_valid(sup)` / `env_valid(sup)`.
- `cert_key(certificado)` → `"iso" | "iatf" | "env"`.
- `load_received(rdb, suppliers, folder="")`: apaga registros sem arquivo
  (`prune_missing_files`), encerra os que o Excel já mostra válidos e soma os
  arquivos encontrados nas pastas (`cert_files.scan_keys`). Devolve o conjunto
  de pares `(rbsno, "iso"|"iatf"|"env")`.
- `build_rows(sup, email, hdb, overdue_days, received=(), risk=())` → lista de
  3 dicts (um por aba). Chaves:

| Chave | Significado |
|---|---|
| `tab` | aba onde aparece: `iso_expiry`, `env_expiry`, `iatf_expiry` |
| `field` | chave usada em histórico/envio (`iso_expiry`, `env_expiry`, `iatf_expiry` ou `combined_cert`) |
| `kind` | `normal`, `blue_iso`, `blue_iatf`, `received` |
| `expiry`, `tier`, `days_left` | situação (`tier == BLUE` nas linhas azuis) |
| `pending` | **será enviado agora** |
| `shown_pending` | só exibição (linha espelho da aba IATF) |
| `notify` | se faz sentido haver lembrete para esta linha |
| `received` / `risk` | flags de pausa |
| `label`, `rbsno`, `pmd`, `name`, `email` | dados de apoio |

Algoritmo ISO × IATF (resumo): a linha de uma aba é **normal** se o próprio
certificado está ativo, ou se está válido e o outro também não está ativo; senão,
se o outro está válido, é **azul**; se nenhum é válido, é a linha **combinada**
(tier `EXPIRED` ou `NO_DATE`), cujo `pending` só existe na aba ISO — a aba IATF
fica com `pending=False` e `shown_pending` espelhando.
Certificado **recebido** (`(rbsno, "iso"|"iatf"|"env")` em `received`): antes
do algoritmo, o tier daquele certificado é anulado (`None` = Ativo) e ele conta
como válido — assim a outra aba vira azul. A linha do próprio certificado vira
`kind="received"` (`received=True`, `pending=False`). Para as demais linhas com
`notify` de ISO/IATF: se o fornecedor está em `risk` (apenas Risk Assessment
**não concluído**) → `risk=True`, `pending=False`.

### `templates.py`
`build_email_body(name, cert_label, expiry_str, days_left, tier, company, language)`.
`tier == NO_DATE` → `_body_missing_pt` (sempre pt). Demais → `_body_pt`
(`_body_en` existe mas não é usado: app só atende BR).

### `email_sender.py`
`send_email(recipient, bcc, subject, html_body, sender="")`: `import
win32com.client` tardio; `mail.BCC`; remetente via `SendUsingAccount` (conta
configurada no Outlook) ou `SentOnBehalfOfName` (fallback).

### `history_db.py`
`HistoryDB`: `get_last_send`, `log_send`, `history(limit)`, `count()`.

### `received_db.py`
`ReceivedDB`: `add` (substitui registro com o mesmo `saved_path`),
`prune_missing_files()`, `all_certs()`, `delete_cert`, `suppressed_groups()`,
`delete_group`, `supplier_count()`; `group_of(certificado)`.

### `risk_db.py`
`RiskDB`: `all`, `add`, `update`, `delete`, `keys()` (RBSNOs **e** nomes em
minúsculo, só de registros **não** `Completed`). Migração automática das colunas
`files` e `evidence` (`ALTER TABLE`). `files` = certificados; `evidence` =
arquivos e links; ambos como texto separado por `|` (links começam com
`http://`/`https://`). Cada Add cria um registro novo (histórico preservado).

### `cert_files.py`
`supplier_folder`, `file_label(cert)` (`ISO 9001`/`IATF 16949`/`ISO 14001`),
`save_file(base, name, rbsno, src, certificado)` (copia e renomeia para
`<Certificado> - <Fornecedor> - <RBSNO>.<ext>`, **sobrescrevendo** o de mesmo
nome), `scan_keys` (lê os prefixos na pasta), `standardize_existing` (renomeia
arquivo antigo para o padrão) e `save_extra` (mantém o nome original, nunca
sobrescreve; usado nas evidências, subpasta `Evidence`). Sanitiza caracteres
inválidos do Windows.

### `notifications.py`
`notify(title, message, log=None)`: tenta `plyer.notification`; qualquer exceção
vira linha de Log (não derruba o app).

### `auto_run.py`
`run_automatic()` replica o envio da GUI usando `build_rows`: carrega config,
Excel, Access, `received`, `risk`; para cada linha `pending` com e-mail, monta
assunto/corpo, envia, registra no histórico e loga. Retorna código de saída.

## 15. Modelo de dados

SQLite único (`tracking.db`):

```sql
notifications(id, rbsno, pmd, supplier_name, cert_type, expiry_date, tier,
              sent_date, recipient_email)
risk_assessment(id, rbsno, pmd, supplier_name, certificate, status,
                added_date, notes, files, evidence)
received_certs(id, rbsno, pmd, supplier_name, certificate, grp, file_name,
               saved_path, received_date)
```

- `notifications.cert_type` ∈ `iso_expiry | env_expiry | iatf_expiry |
  combined_cert`; `expiry_date` é `str(date)` (ou `"None"` no caso em branco).
- Cada classe abre/fecha sua própria conexão e cria sua tabela no construtor
  (`CREATE TABLE IF NOT EXISTS`).

## 16. Fluxo de execução

**Análise (GUI)** — `Load & Analyze`:
1. `_load_and_analyze` valida o arquivo e dispara `_analyze_worker` em thread.
2. Worker: define `WARNING_THRESHOLDS`, `read_suppliers(..., country_filter="BR")`,
   `load_emails` (se o Access estiver configurado; avisos vão ao Log),
   `load_received` (com a pasta de certificados), `RiskDB.keys()`, e `build_rows` por fornecedor.
3. Emite `analysis_finished(suppliers, rows)` → `_on_analysis_finished` (thread
   da UI) preenche tabelas, cartões, contador, gráficos e Follow-up.

**Envio** — `Send Pending Reminders`: confirma; `_send_worker` em thread; para
cada linha `pending` com e-mail: `build_email_body` → `send_email` →
`HistoryDB.log_send` → signal de progresso; no fim, `send_finished` e
notificação do Windows.

## 17. Interface (Qt)

- **`MainWindow`**: cabeçalho + `QHBoxLayout` com **sidebar** (`QFrame#Sidebar`,
  `QButtonGroup` exclusivo) e `QStackedWidget` com as 6 páginas
  (`nav_pages = [panel, risk, reports, history, log, settings]`).
  `_navigate(key)` troca a página; `_toggle_sidebar()` alterna largura
  (215/64 px), remove o texto e seta a propriedade `collapsed` para o QSS
  centralizar os ícones.
- **Ícones**: desenhados com `QPainter` em `make_nav_icon(kind)` (panel, risk,
  reports, history, log, settings, menu) — sem arquivos externos.
- **Tabelas**: `QTableWidget` com `SortItem` (`ui/common.py`), que ordena por
  `sort_key` em vez do texto. Coluna Status: chave
  `(texto sem dígitos, dias)`; RBSNO/PMD: `numeric_or_text_key`.
- **Cores**: `ROW_COLORS` (active, expiring, expired, no_date, blue, risk) em
  `main_window.py`; paleta e QSS em `ui/theme.py` (`GREEN, AMBER, RED, BLUE,
  PURPLE`…).
- **Gráficos** (`ui/pages.py`): `DonutChart` e `BarChart` próprios com
  `QPainter` (sem QtCharts, para reduzir dependências de DLL). `ReportsPage.set_data`
  e `set_follow`.
- **QSS**: seletores por `objectName` (`#SidebarButton`, `#StatCard`,
  `#PendingCounter`, `#TimeSavedLabel`…). Labels que ficam sobre cartões
  precisam de `background: transparent` (o seletor global `QWidget` pinta tudo).
- **Threads**: use sempre os signals de `Worker`; não toque em widgets fora da
  thread da UI.

## 18. Como estender o app

**Novo tipo de certificado**
1. `config.py`: chave `excel_col_<novo>`.
2. `excel_reader.py`: `DEFAULT_COLUMNS`, `FIELD_LABELS`, `SupplierRecord`,
   `_extract_suppliers`.
3. `supplier_rows.py`: entradas em `CERT_TYPES`, `CERT_TYPE_LABELS_PT`,
   `CERT_ISO_TAG`, `CERT_LABEL_EN` e a lógica em `build_rows`.
4. UI: campo em Settings (`_add_field`), `auto_run.py` (`excel_columns`),
   relatórios (`REPORT_TYPES`) se couber.

**Nova faixa/regra de envio:** `reminder_engine.py` (`needs_notification`) — a
GUI e o `--auto` herdam.

**Novo texto de e-mail:** nova função em `templates.py` + ramo em
`build_email_body` + opção em `EmailPreviewDialog.TIER_OPTIONS`.

**Nova página:** classe em `ui/pages.py`, adicionar ao `QStackedWidget` e a
`nav_pages`/`nav_labels`, e um ícone em `make_nav_icon`.

**Nova coluna da tabela:** `TABLE_HEADERS`, `CENTERED_COLUMNS` e a construção da
linha em `_on_analysis_finished`.

## 19. Testes e depuração

Não há suíte automatizada. A validação foi feita com **Qt headless**:

```bash
pip install PySide6 openpyxl plyer
# pyodbc precisa de unixodbc no Linux; para testar a UI sem Access, use um stub:
#   sys.modules["pyodbc"] = types.ModuleType("pyodbc")
QT_QPA_PLATFORM=offscreen python teste.py
```

Padrão usado: criar `QApplication`, instanciar `MainWindow`, apontar `cfg`
para um Excel sintético e um `.db` temporário, chamar `_load_and_analyze()`,
aguardar ~1,5 s (`QTimer.singleShot`) e inspecionar as tabelas/cartões ou
salvar `w.grab().save("x.png")` para conferir o visual. Para a regra de negócio
pura, chame `supplier_rows.build_rows` direto com `SupplierRecord` fabricados.

Dicas:
- Em testes, evite `QMessageBox` modal (substitua `QMessageBox.critical` etc.).
- `py_compile` em todos os módulos antes de empacotar.
- Antes de gerar o zip, remova `__pycache__`, `tracking*.db/json/txt`.

## 20. Decisões de projeto e limitações conhecidas

- **Só BR:** filtro por Country Code na leitura; e-mails só em português.
- **Envio só no Windows** (Outlook COM). Em outros SOs, só a interface roda.
- **Reenvio de "Missing date"** segue o mesmo intervalo dos vencidos (15 dias).
- **Certificado expirando + outro ativo → sem aviso** (decisão de negócio:
  um repõe o outro).
- **Risk Assessment em aberto pausa a cobrança; Completed não** (fica como
  histórico).
- **Certificado recebido = Ativo enquanto o arquivo existir:** o app não lê a
  data de validade do arquivo; a identificação é pelo **nome** do arquivo.
- **Missing date (Follow-up)** é calculado por **subtração de contagens**
  (`missing_total − risk`, onde `missing_total` já exclui os recebidos), não por interseção de fornecedores; se
  um fornecedor em Risk Assessment não estiver sem data, o número ainda cai.
- **Sem confirmação de entrega/resposta:** o app não lê a caixa de entrada; o
  "recebido" é registrado manualmente.
- **Histórico por RBSNO + certificado + data de validade:** mudar a data no
  Excel inicia novo ciclo de avisos.
- **Conexões SQLite curtas:** cada operação abre/fecha; suficiente para o volume
  (centenas de fornecedores).
- **Gráficos próprios** em vez de QtCharts, para evitar problemas de DLL já
  vistos em ambientes corporativos.
- PySide6 **fixado em 6.7.2** (outras versões causaram `DLL load failed`
  em alguns ambientes Windows).

## 21. Histórico de mudanças

- **Tracking Certificate:** renomeado a partir de RenovaCert; arquivos
  `renovacert_*` → `tracking_*` (`tracking.db`), migração automática.
- Reescrita Tkinter → **PySide6**, interface em inglês.
- Sub-abas por certificado; cartões dinâmicos; ordenação por coluna.
- Assunto do e-mail termina com o **RBSNO** (antes, nome do fornecedor); **Shared Database**: local do `tracking.db` configurável (banco único em rede para vários PCs).
- Filtro **BR**; assunto com tag de certificado; coluna **Reminder**;
  "Missing date".
- **Send History** com total de envios e **Time Saved** (5 min/envio).
- **Sidebar** recolhível com ícones; Settings na base.
- **IATF 16949** + substituição cruzada com ISO 9001; e-mail combinado.
- **BCC**, **remetente configurável**, novos padrões de fábrica.
- **Log** com hora, persistente em arquivo; **notificações** (Plyer).
- **E-mail de certificado em branco**; **Risk Assessment** (CRUD, filtros,
  anexos); **Register Received Certificate**; **Reports** (Quality,
  Environmental, Follow-up).
- Último Excel lembrado; ordenação alfabética de Status; Status centralizado.
- **Time Saved** movido para Reports; **campo de pesquisa** no Panel
  (Supplier/RBSNO/PMD); fundo dos rótulos do Settings corrigido.
- **Certificado recebido conta como Ativo** (por certificado, com azul cruzado
  ISO/IATF e Ambiental independente); arquivos renomeados no padrão
  `<Certificado> - <Fornecedor> - <RBSNO>`; conferência da pasta a cada análise
  (arquivo apagado remove o registro); mesmo certificado substitui o anterior.
- **Risk Assessment:** campos **Certificate** e **Evidence** (arquivos e links),
  exclusão física só ao salvar, histórico preservado e Completed não pausa mais.
