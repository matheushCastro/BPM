
#   Preciso ler arquivo linha por linha
#   Mandar este texto para a Ia e me retornar o xml
#   Salvar e abrir no bizagi ou camunda
# =====================================================================

import re       # expressões regulares (usada para limpar o texto da IA)
import sys      # permite receber o caminho do arquivo pelo terminal
import time     # usada para esperar entre as tentativas
from pathlib import Path

from docx import Document          # lê arquivos .docx
from google import genai           # conversa com o Gemini
from google.genai import types     # configurações da chamada à IA
from lxml import etree             # lê e confere XML




#modelos do gemini para ir testando caso algum estiver ocupado vai passando para o outro

MODELOS = ["gemini-3.8-flash", "gemini-flash-latest", "gemini-flash-lite-latest"]

# O PROMPT é o "manual de instruções" que a IA recebe. É aqui que você
# controla o que ela faz. Se o desenho sair errado, é este texto que se ajusta.
PROMPT = """Converta o POP em português para um arquivo BPMN 2.0 completo.
Responda APENAS com o XML, sem texto extra e sem ```.

Conteúdo do XML:
- <definitions> com os namespaces: BPMN (MODEL) como padrão, bpmndi (DI), dc (DC) e di (DD/DI).
- <collaboration> com um <participant> que aponta para o <process>.
- <process> com um <laneSet>: uma <lane> por responsável (Solicitante, Secretaria etc.),
  cada uma listando seus elementos em <flowNodeRef>.
- Elementos: startEvent, task (ou userTask/sendTask/manualTask quando fizer sentido),
  exclusiveGateway e endEvent, ligados por sequenceFlow.
- Nomes das tarefas: verbo no infinitivo + substantivo, até 4 palavras.
- Toda decisão ("se... senão...") é um exclusiveGateway com nome em forma de pergunta
  e DUAS saídas, cada sequenceFlow com name="Sim" ou name="Não".
- Um endEvent para cada desfecho do processo.
- Ignore objetivo, escopo e observações gerais.

Desenho (obrigatório, dentro de <bpmndi:BPMNDiagram> e <bpmndi:BPMNPlane>):
- Um BPMNShape para o participant e para cada lane (isHorizontal="true"), um para cada
  elemento e um BPMNEdge (com di:waypoint) para CADA sequenceFlow.
- Piscina em x=40, y=40. Cada raia tem 120 de altura, uma abaixo da outra, em x=70.
- Cada elemento ocupa a sua PRÓPRIA coluna, na ordem do fluxo: x = 110 + índice * 170.
  Assim nenhum elemento se sobrepõe a outro.
- A altura vem da raia: centralize o elemento verticalmente na faixa da sua raia.
- Tamanhos: eventos 36x36, gateways 50x50, tarefas 100x80.
- Setas saem do meio da borda direita da origem e chegam no meio da borda esquerda
  do destino. Se a altura for diferente, use 4 waypoints formando um degrau."""



#Endereços de referencias pra quando ler o xml o camunda reconhecer
BPMN_MODEL = "http://www.omg.org/spec/BPMN/20100524/MODEL"
NAMESPACES = {
    "bpmndi": "http://www.omg.org/spec/BPMN/20100524/DI",
    "dc": "http://www.omg.org/spec/DD/20100524/DC",
    "di": "http://www.omg.org/spec/DD/20100524/DI",
}

def corrigir_namespaces(xml):
    """Troca os endereços declarados no XML pelos corretos."""
    for prefixo, endereco in NAMESPACES.items():
        xml = re.sub(rf'xmlns:{prefixo}="[^"]*"', f'xmlns:{prefixo}="{endereco}"', xml)
    # o endereço padrão (sem prefixo) é o do processo
    xml = re.sub(r'xmlns="[^"]*"', f'xmlns="{BPMN_MODEL}"', xml, count=1)
    return xml

#ler o pop de acordo com o passo a passo que se tem nele 
def ler_pop(caminho):
    """Abre o .docx e devolve uma lista com o texto de cada linha."""
    caminho = Path(caminho)

    # Confere se o arquivo é válido antes de tentar abrir
    if caminho.suffix.lower() != ".docx":
        raise ValueError("O arquivo deve estar no formato DOCX.")
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {caminho}")

    documento = Document(caminho)
    linhas = []

    # Parágrafos normais. O " ".join(texto.split()) tira espaços repetidos.
    for paragrafo in documento.paragraphs:
        texto = " ".join(paragrafo.text.split())
        if texto:  # ignora linhas vazias
            linhas.append(texto)

    # Tabelas (como a tabela de Responsáveis). Cada linha da tabela vira
    # um texto com as células separadas por " | ".
    for tabela in documento.tables:
        for linha in tabela.rows:
            celulas = [" ".join(celula.text.split()) for celula in linha.cells]
            linhas.append(" | ".join(celulas))

    return linhas


# ---------------------------------------------------------------------
# PASSO 2: PEDIR O XML PARA A IA
# ---------------------------------------------------------------------
def gerar_xml(linhas):
    """Envia o texto do POP para o Gemini e devolve o XML (ainda sem conferir)."""
    # timeout=90000 -> desiste se a IA não responder em 90 segundos
    cliente = genai.Client(http_options=types.HttpOptions(timeout=90000))
    espera = 5  # segundos de espera após uma falha (dobra a cada vez)

    # Tenta cada modelo da lista, até 3 vezes cada
    for modelo in MODELOS:
        for tentativa in range(1, 4):
            try:
                print(f"Tentando {modelo} ({tentativa}/3)...")
                resposta = cliente.models.generate_content(
                    model=modelo,
                    contents="\n".join(linhas),          # o texto do POP
                    config=types.GenerateContentConfig(
                        system_instruction=PROMPT,       # as instruções
                        temperature=0,                   # 0 = respostas mais estáveis
                    ),
                )
                return resposta.text  # deu certo: devolve o XML

            except Exception as erro:
                texto = str(erro).lower()

                # vai dar erro caso o modelo da minha chave não for esse e ela testa o proximo
                if "404" in texto:
                    print(f"{modelo} não disponível, passando para o próximo.")
                    break

                # Erros temporários: modelo ocupado, limite, demora -> espera e repete
                temporario = any(c in texto for c in ("503", "504", "429", "timed out", "timeout"))
                if not temporario:
                    raise  # se a chave que eu tenho estiver errada

                print(f"Ocupado, esperando {espera}s...")
                time.sleep(espera)
                espera = min(espera * 2, 30)  # 5, 10, 20, 30, 30...

    raise RuntimeError("Nenhum modelo respondeu agora. Tente de novo em alguns minutos.")


# ---------------------------------------------------------------------
# PASSO 3: CONFERIR O XML
# ---------------------------------------------------------------------
def limpar_e_validar(xml):
    """Limpa o texto da IA e confere se o XML está completo.
    Se algo estiver errado, levanta um ValueError (o converter tenta de novo)."""

    # Às vezes a IA coloca ```xml ... ``` em volta. Aqui isso é removido.
    xml = re.sub(r"^```(?:xml)?\s*|\s*```$", "", (xml or "").strip())
    xml = corrigir_namespaces(xml)  

    # 1) O XML é bem formado? (tags abertas e fechadas corretamente)
    try:
        raiz = etree.fromstring(xml.encode("utf-8"))
    except etree.XMLSyntaxError as erro:
        raise ValueError(f"XML inválido: {erro}")

    # 2) Tem o processo e o desenho? (sem o desenho o Camunda abre em branco)
    if not raiz.findall(".//{*}process"):
        raise ValueError("O XML não tem <process>.")
    if not raiz.findall(".//{*}BPMNDiagram"):
        raise ValueError("O XML não tem o desenho (bpmndi).")

    # Junta todos os ids do arquivo para comparar abaixo
    ids = {e.get("id") for e in raiz.iter() if e.get("id")}

    # 3) Toda seta liga elementos que realmente existem?
    fluxos = raiz.findall(".//{*}sequenceFlow")
    for fluxo in fluxos:
        if fluxo.get("sourceRef") not in ids or fluxo.get("targetRef") not in ids:
            raise ValueError("Uma seta liga a um elemento que não existe.")

    # 4) Todo desenho aponta para algo que existe?
    desenhados = set()
    for item in raiz.iter("{*}BPMNShape", "{*}BPMNEdge"):
        if item.get("bpmnElement") not in ids:
            raise ValueError("Um desenho aponta para um elemento que não existe.")
        desenhados.add(item.get("bpmnElement"))

    # 5) Toda seta tem o seu desenho? (senão ela some no Camunda)
    for fluxo in fluxos:
        if fluxo.get("id") not in desenhados:
            raise ValueError("Uma seta ficou sem desenho.")

    #verifica se o desenho esta de acordo com endereço certo, os que coloquei lá em cima
    for item in raiz.iter("{*}BPMNDiagram", "{*}BPMNPlane", "{*}BPMNShape", "{*}BPMNEdge"):
        if etree.QName(item).namespace != NAMESPACES["bpmndi"]:
            raise ValueError("O desenho está com o namespace bpmndi errado.")
    for item in raiz.iter("{*}Bounds"):
        if etree.QName(item).namespace != NAMESPACES["dc"]:
            raise ValueError("Um <Bounds> está com o namespace dc errado.")
    for item in raiz.iter("{*}waypoint"):
        if etree.QName(item).namespace != NAMESPACES["di"]:
            raise ValueError("Um <waypoint> está com o namespace di errado.")

    return xml


# ---------------------------------------------------------------------
# PASSO 4: JUNTAR TUDO E SALVAR
# ---------------------------------------------------------------------
def converter(caminho_entrada):
    """Faz o processo todo: lê o POP, pede o XML, confere e salva o .bpmn."""
    entrada = Path(caminho_entrada)

    print("Lendo o POP...")
    linhas = ler_pop(entrada)
    print(f"{len(linhas)} linhas lidas. Pedindo o XML para a IA...")

    # A IA pode errar de vez em quando (esquecer uma seta, por exemplo).
    # vamos pedir 3 vezes pois tem risco da IA erra , meio que ela vai se corigindo se necessario
    for tentativa in range(1, 4):
        try:
            xml = limpar_e_validar(gerar_xml(linhas))
            break  # o XML passou na conferência: sai do laço
        except ValueError as erro:
            print(f"XML recusado ({erro}). Pedindo de novo ({tentativa}/3)...")
    else:
        # o "else" do for só roda se o laço terminou sem nenhum "break"
        raise RuntimeError("A IA não gerou um XML válido em 3 tentativas.")

    # Salva ao lado do .docx, com o mesmo nome e a extensão .bpmn
    saida = entrada.with_suffix(".bpmn")
    saida.write_text(xml, encoding="utf-8")
    return saida



if __name__ == "__main__":
 
    arquivo = sys.argv[1] if len(sys.argv) > 1 else input("Coloque o caminho do docx: ").strip('"')

    try:
        print(f"BPMN criado: {converter(arquivo)}")
    except Exception as erro:
               print(f"Erro: {erro}")