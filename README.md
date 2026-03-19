 # Nesting 2D v2.0

Ferramenta web local para otimização de corte de chapas.
Interface com três painéis: configuração, visualização SVG interativa e lista de layouts.

---

## Instalação

### 1. Python 3.10+
https://www.python.org/downloads/
Marque **"Add Python to PATH"**.

### 2. Dependências
```
pip install -r requirements.txt
```

---

## Uso

```
python app.py
```
Abre automaticamente em `http://localhost:5000`.

---

## Interface

**Painel esquerdo — Configuração**
- Modo **Retangular**: defina peças diretamente (nome, L, A, quantidade, pode girar)
- Modo **DXF**: aponte a pasta com os arquivos `.dxf`
- Parâmetros de chapa, gap, margens, material e custo

**Canvas central**
- Visualização SVG do layout da chapa selecionada
- Zoom com scroll do mouse ou botões +/−
- Pan com clique e arrasto
- Dimensões cotadas em cada peça e nas bordas da chapa
- Tooltip com detalhes ao passar o mouse

**Painel direito — Layouts**
- KPIs: chapas necessárias, aproveitamento médio, peças alocadas
- Lista de chapas clicável — clique para ver o layout
- Tabela de peças: solicitadas vs. alocadas por tipo

**Exportar**
- Botão "Exportar .xlsx" no cabeçalho
- Gera relatório com aba RESUMO + uma aba por chapa (imagem do layout + tabela)

---

## Modos de algoritmo

**Retangular (Guillotine)**
- Para peças retangulares
- Algoritmo de corte guilhotina com heurística BSSF
- Mais rápido e preciso para este tipo de peça

**DXF (Bottom-Left Fill)**
- Para geometrias arbitrárias vindas de arquivos DXF
- Suporta: LINE, LWPOLYLINE, ARC, CIRCLE, SPLINE
- Requer Shapely e ezdxf instalados

---

## Dicas para DXF

- Uma peça por arquivo
- Contorno fechado, sem cotas ou hatch
- Unidade: milímetros
- Exportar do Inventor com geometria explodida (sem blocos)

---

## Estrutura

```
nesting2/
├── app.py           ← servidor Flask
├── engine.py        ← algoritmos (guillotine + BL-fill)
├── export.py        ← relatório .xlsx
├── requirements.txt
├── templates/
│   └── index.html   ← interface
└── README.md
```
