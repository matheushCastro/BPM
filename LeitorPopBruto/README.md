# Leitor de POP para BPMN

Ferramenta de linha de comando que extrai etapas de um POP em PDF ou Word
(`.docx`) e gera um arquivo BPMN 2.0 (`.bpmn`) com eventos, tarefas, fluxos e
layout BPMN-DI para abrir em ferramentas de modelagem compatíveis.

## Requisitos

- Python 3.10 ou superior
- Para arquivos PDF, instale a dependência:

```powershell
python -m pip install -r requirements.txt
```

Arquivos DOCX são lidos com a biblioteca padrão do Python, sem dependências
adicionais. PDFs digitalizados (imagem) precisam passar por OCR antes da leitura.

## Uso

```powershell
python pop_to_bpmn.py "C:\caminho\meu-pop.docx"
python pop_to_bpmn.py "C:\caminho\meu-pop.pdf" --saida "C:\caminho\processo.bpmn"
python pop_to_bpmn.py "C:\caminho\meu-pop.docx" --nome "Atendimento ao cliente"
```

Por padrão, o arquivo BPMN é criado ao lado do POP, com o mesmo nome. Confira as
etapas reconhecidas antes de exportar:

```powershell
python pop_to_bpmn.py "C:\caminho\meu-pop.docx" --preview
```

## Como o leitor interpreta o POP

O leitor procura uma seção chamada “Procedimento”, “Descrição do procedimento”,
“Execução”, “Passo a passo” ou equivalente. Se encontrar uma, ignora o conteúdo
anterior e termina a leitura em seções como “Referências” ou “Registros”.
Etapas numeradas são priorizadas; sem numeração, linhas de texto são tratadas
como etapas. Revise o `--preview`: decisões, responsáveis e ramificações não são
inferidos automaticamente; o texto das etapas é preservado em tarefas BPMN
sequenciais para revisão e ajuste na ferramenta de mapeamento.

## Testes

```powershell
python -m unittest
```
