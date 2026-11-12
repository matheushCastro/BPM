import re
import os
from docx import Document            
from xml.sax.saxutils import quoteattr, escape

arquivo = input("Cole o caminho do documento (.docx): ").strip().strip('"').strip("'")
if not os.path.exists(arquivo):
    print(f"Erro: arquivo '{arquivo}' não encontrado.")
    exit(1)

PADRAO = re.compile(r'^(\d+)\s*[.)]\s*(.+)')   #    um parágrafo que começa com o próximo número esperado (1,2 )

# 1) Ler as etapas. Um passo é:
#    um item de lista do Word (marcador ou numeração automática), OU
#    um padrão decidido que apos o decidido um pragrafo tem que conferir se vai ter o
#    próximo numero e uma quebra de linha 
#    Títulos (Heading) antes dos passos são ignorados; um título depois dos passos encerra a leitura.
etapas = []                                     # texto completo de cada etapa
for p in Document(arquivo).paragraphs:
    linha = p.text.strip()
    if not linha:
        continue
#testando git
    pPr = p._p.pPr
    estilo = pPr.pStyle.val.lower() if (pPr is not None and pPr.pStyle is not None) else ""
    eh_titulo = estilo.startswith(("heading", "ttulo", "título", "title"))
    eh_lista = pPr is not None and pPr.numPr is not None

    if eh_titulo:
        if etapas:
            break                               # novo título depois dos passos = fim do procedimento
        continue                                # título antes dos passos: ignora

    m = PADRAO.match(linha)
    if eh_lista:
        etapas.append(linha)                    # item de lista = um passo
    elif m and int(m.group(1)) == len(etapas) + 1:
        etapas.append(m.group(2).strip())       # número proximo esperado (1,2,3)
    elif etapas:
        etapas[-1] += " " + linha               # continuação do passo anterior
    # antes do primeiro passo (título, objetivo...) cai aqui e é ignorado

if not etapas:
    print("Nenhuma etapa encontrada (use uma lista do Word ou comece os passos com '1.').")
    exit(1)


# 2) Rótulo curto para a caixa (padrão BPMN: ação curta, não a frase inteira)
def rotulo(texto, max_palavras=6):
    parte = re.split(r'[,;:.]', texto)[0]       # corta na 1ª vírgula, ; : ou .
    palavras = parte.split()[:max_palavras]     # no máximo 6 palavras
    soltas = {"a", "o", "as", "os", "de", "da", "do", "das", "dos", "no", "na",
              "nos", "nas", "em", "e", "ou", "para", "por", "com", "ao", "um", "uma"}
    while len(palavras) > 1 and palavras[-1].lower() in soltas:
        palavras.pop()                          # tira palavra solta no final ("...sala no")
    return " ".join(palavras).rstrip(".")


# Como o mapeamento e bruto ele so vai ligando o inicio com as atividades ate o final 0 --== 0
# e vai decidindo tambem seu tamanho de inicio e atividades e o fim
nome_processo = os.path.splitext(os.path.basename(arquivo))[0]
nos = [("Inicio", "startEvent", "Início", 36, "")]
nos += [(f"Etapa_{i}", "task", rotulo(t), 120, t) for i, t in enumerate(etapas, 1)]
nos += [("Fim", "endEvent", "Fim", 36, "")]

processo, formas, setas = "", "", ""
x = 150
for i, (id_, tag, nome, w, completo) in enumerate(nos):
    h = 80 if tag == "task" else 36
    entra = f"F_{nos[i - 1][0]}_{id_}" if i > 0 else None
    sai = f"F_{id_}_{nos[i + 1][0]}" if i < len(nos) - 1 else None

    filhos = ""
    if completo:   # texto completo vai na documentação (aparece no painel do Camunda)
        filhos += f'<bpmn:documentation>{escape(completo)}</bpmn:documentation>'
    if entra:
        filhos += f'<bpmn:incoming>{entra}</bpmn:incoming>'
    if sai:
        filhos += f'<bpmn:outgoing>{sai}</bpmn:outgoing>'
    processo += f'<bpmn:{tag} id="{id_}" name={quoteattr(nome)}>{filhos}</bpmn:{tag}>\n'

    formas += (f'<bpmndi:BPMNShape id="S_{id_}" bpmnElement="{id_}">'
               f'<dc:Bounds x="{x}" y="{200 - h // 2}" width="{w}" height="{h}" /></bpmndi:BPMNShape>\n')
    if entra:
        processo += f'<bpmn:sequenceFlow id="{entra}" sourceRef="{nos[i - 1][0]}" targetRef="{id_}" />\n'
        setas += (f'<bpmndi:BPMNEdge id="E_{entra}" bpmnElement="{entra}">'
                  f'<di:waypoint x="{x - 50}" y="200" /><di:waypoint x="{x}" y="200" /></bpmndi:BPMNEdge>\n')
    x += w + 50

largura_pool = x - 30   # a piscina (pool) envolve todo o fluxo
refs = "".join(f"<bpmn:flowNodeRef>{n[0]}</bpmn:flowNodeRef>" for n in nos)

# 4) Gerar o XML: pool com uma raia (lane) contendo todos os nós, como o Bizagi organiza o modelo
xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<bpmn:definitions xmlns:bpmn="http://www.omg.org/spec/BPMN/20100524/MODEL"
                  xmlns:bpmndi="http://www.omg.org/spec/BPMN/20100524/DI"
                  xmlns:dc="http://www.omg.org/spec/DD/20100524/DC"
                  xmlns:di="http://www.omg.org/spec/DD/20100524/DI"
                  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                  id="Definitions_1"
                  targetNamespace="http://bpmn.io/schema/bpmn">
<bpmn:collaboration id="Collaboration_1">
<bpmn:participant id="Pool_1" name={quoteattr(nome_processo)} processRef="Process_1" />
</bpmn:collaboration>
<bpmn:process id="Process_1" name={quoteattr(nome_processo)} isExecutable="false">
<bpmn:laneSet id="LaneSet_1"><bpmn:lane id="Lane_1" name="Atividades">{refs}</bpmn:lane></bpmn:laneSet>
{processo}</bpmn:process>
<bpmndi:BPMNDiagram id="D_1"><bpmndi:BPMNPlane id="P_1" bpmnElement="Collaboration_1">
<bpmndi:BPMNShape id="S_Pool_1" bpmnElement="Pool_1" isHorizontal="true">
<dc:Bounds x="30" y="100" width="{largura_pool}" height="200" /></bpmndi:BPMNShape>
<bpmndi:BPMNShape id="S_Lane_1" bpmnElement="Lane_1" isHorizontal="true">
<dc:Bounds x="60" y="100" width="{largura_pool - 30}" height="200" /></bpmndi:BPMNShape>
{formas}{setas}</bpmndi:BPMNPlane></bpmndi:BPMNDiagram>
</bpmn:definitions>
'''

saida = arquivo.rsplit(".", 1)[0] + ".bpmn"
with open(saida, "w", encoding="utf-8") as f:
    f.write(xml)
print(f"Sucesso: {len(etapas)} etapas -> {saida}")