[README.md](https://github.com/user-attachments/files/33387323/README.md)
# Gerador de Etiquetas 

Aplicativo em Streamlit para gerar etiquetas a partir de uma planilha com os SIDs selecionados e de uma planilha `QUIMICO_SIMPLES` com os dados de identificação das amostras.

## O que o aplicativo faz

Para cada célula preenchida na planilha `ETIQUETAS`, o aplicativo cria uma etiqueta com três linhas:

1. Parâmetro e SID, por exemplo `m 003`.
2. OS e talhão Agrorobótica, por exemplo `OS331|A1`.
3. Remessa e tipo de análise, por exemplo `R1|F`.

O aplicativo cruza cada etiqueta com a planilha `QUIMICO_SIMPLES` usando a OS, a remessa, o tipo de análise e o SID. Assim, SIDs iguais de OS diferentes não são confundidos.

## Arquivos do projeto

Mantenha estes arquivos na pasta do projeto ou na raiz do repositório do GitHub:

```text
Gerador_Etiquetas_ATVOS.py
requirements.txt
README.md
```

Conteúdo esperado de `requirements.txt`:

```text
streamlit
openpyxl
reportlab
```

As planilhas usadas para gerar etiquetas são carregadas na interface. Não é necessário armazená-las no repositório.

## Planilhas de entrada

### Planilha ETIQUETAS

- A primeira aba da planilha é lida pelo aplicativo.
- A primeira linha deve conter as colunas de parâmetros: `k`, `m`, `u`, `s`, `p`, `c`, `t`, `r` e `CHN`.
- Cada célula preenchida deve seguir o padrão `parâmetro SID|OS|remessa|tipo`, por exemplo `m 003|331|R1|F`.
- O parâmetro escrito na etiqueta deve corresponder à coluna em que ela está.
- Células vazias são ignoradas. Cada célula preenchida gera uma etiqueta.

### Planilha QUIMICO_SIMPLES

A primeira aba deve conter estas colunas, com os nomes escritos exatamente assim:

- `Ordem_de_Serviço`
- `Talhão_AG`
- `SID`

A coluna `Ordem_de_Serviço` deve seguir o padrão `OS_331_A1_R1_F`, em que as partes identificam a OS, o talhão, a remessa e o tipo de análise. O aplicativo usa `Talhão_AG` como origem do talhão impresso na etiqueta.

## Como usar o aplicativo

1. Abra o aplicativo.
2. Em **Planilha ETIQUETAS**, selecione a planilha com os parâmetros e SIDs.
3. Em **Planilha QUIMICO_SIMPLES**, selecione a planilha com a OS, o `Talhão_AG` e o SID.
4. Clique em **Gerar etiquetas**.
5. Se a validação for concluída, clique em **Baixar PDFs em ZIP**.

O ZIP contém um PDF para cada combinação de parâmetro e tipo de análise encontrada. Cada página comporta até quatro etiquetas lado a lado. As três linhas têm espaçamento vertical de 1 mm.

## Executar no computador

Instale o Python e abra o PowerShell na pasta que contém o código. Execute:

```powershell
python -m pip install -r requirements.txt
streamlit run Gerador_Etiquetas_ATVOS.py
```

O Streamlit abrirá o aplicativo no navegador. O endereço local normalmente começa com `http://localhost:8501` e só é acessível no computador em que o aplicativo está rodando.

## Publicar no Streamlit Community Cloud

1. Envie `Gerador_Etiquetas_ATVOS.py`, `requirements.txt` e `README.md` para o repositório do GitHub.
2. Acesse [share.streamlit.io](https://share.streamlit.io) e entre com a conta do GitHub que tem acesso ao repositório.
3. Clique em **Create app** e selecione o repositório, a branch e o arquivo `Gerador_Etiquetas_ATVOS.py`.
4. Clique em **Deploy**. Depois que o aplicativo iniciar, copie o endereço terminado em `streamlit.app`.

Um repositório privado gera um app privado por padrão. Para que outras pessoas o acessem, adicione-as como visualizadoras do app. Um app público pode ser acessado por qualquer pessoa com o link. Verifique as regras da empresa antes de enviar planilhas de trabalho para um aplicativo hospedado na nuvem.

## Mensagens de validação

- **Selecione as duas planilhas:** os dois arquivos são obrigatórios.
- **Colunas ausentes:** confira os nomes dos cabeçalhos na primeira linha da aba ativa.
- **Formato inválido na planilha ETIQUETAS:** confira se a célula segue o padrão `m 003|331|R1|F` e se está na coluna do parâmetro correspondente.
- **Não encontrei Talhão_AG:** confira se a OS, a remessa, o tipo e o SID da etiqueta existem juntos na planilha `QUIMICO_SIMPLES`.
- **Mais de um Talhão_AG para a mesma chave:** revise os dados da planilha química para a combinação informada de OS, remessa, tipo e SID.

## Observações

- Não envie planilhas operacionais ao GitHub junto com o código. Elas devem ser escolhidas pelo usuário na tela do aplicativo.
- Se o aplicativo estiver hospedado no Streamlit Community Cloud, as planilhas selecionadas são enviadas para o serviço que executa o app. Confirme que esse uso está autorizado para os dados da empresa.
- O aplicativo usa Times New Roman quando as fontes estão disponíveis no Windows. Em ambientes sem essas fontes, usa Helvetica.
