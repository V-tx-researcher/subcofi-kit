# Kit SubCOFI – gestão de demandas, ações e relatórios

Ferramenta para Subcomissões de Orientação e Fiscalização (SubCOFI) do CRESS-SP.
**Cada Seccional usa a sua própria cópia da planilha, no próprio Google Drive.** Este repositório tem só a mecânica: nenhum dado de Seccional fica aqui.

## O que tem

| Arquivo | Para quê |
|---|---|
| `SubCOFI_Planilha_Modelo.xlsx` | Planilha em branco: Demandas, Ações e eventos, Painel, Listas, Equipe, Municípios |
| `SubCOFI_Script.gs` | Menu SubCOFI (relatório por período, pauta, mapas) e formulário de relatório de atividade |
| `index.html` | Página pública do formulário (este site). Envia os dados para a planilha da Seccional |

## Instalar na sua Seccional (uma vez)

1. No Google Drive, envie `SubCOFI_Planilha_Modelo.xlsx` e use **Arquivo › Salvar como Planilhas Google**.
2. Preencha as abas **Equipe** e **Municípios** (localidades da sua Seccional) e, na aba **Listas**: e-mail do CRESS-SP (E2) e **código de acesso** (G2 – troque o valor padrão).
3. **Extensões › Apps Script**: apague o conteúdo, cole `SubCOFI_Script.gs` e salve.
4. **Implantar › Nova implantação › App da Web**: executar como *Eu*, acesso *Qualquer pessoa*. Autorize e copie o **código da implantação** (começa com `AKfy…`).
5. O link do formulário da sua Seccional é:
   `https://v-tx-researcher.github.io/subcofi-kit/?s=CODIGO_DA_IMPLANTACAO`
   Cole esse link em **Listas › F2** e envie só para a sua equipe, junto com o código de acesso.

## Privacidade

- Os registros, relatórios (.docx) e fotos ficam **só no Drive da Seccional**.
- Sem o código de acesso, a página não carrega listas nem grava nada.
- O formulário **não envia e-mail**: ele monta o texto para a pessoa copiar e enviar da própria conta.
- Não registre nomes de pessoas atendidas, denunciantes ou dados sigilosos.

Mapas por município: disponíveis para as localidades com contorno no script (região de São José do Rio Preto). Nas demais Seccionais, painel e relatórios funcionam normalmente, com o mapa em branco.
