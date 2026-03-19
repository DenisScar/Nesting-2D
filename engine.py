"""
engine.py  —  Motor de nesting 2D
Suporta dois modos:
  - 'rect'  : peças retangulares, algoritmo guillotine + BL-fill
  - 'poly'  : polígonos arbitrários via DXF, algoritmo BL-fill com Shapely
"""

import math
import os
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


# ── Paleta de cores ────────────────────────────────────────────
PALETA = [
    '#4472C4','#ED7D31','#70AD47','#FF0000',
    '#7030A0','#00B0F0','#FFC000','#FF00FF',
    '#00B050','#C00000','#5B9BD5','#92D050',
    '#375623','#833C00','#203864','#843C0C',
]


@dataclass
class Peca:
    id:         str
    nome:       str
    largura:    float         # mm — bounding box
    altura:     float         # mm — bounding box
    area_real:  float         # mm² — área real (com furos descontados)
    quantidade: int
    rotacao:    bool = True   # permite girar 90°
    cor:        str  = '#4472C4'
    polygon:    object = None  # Shapely Polygon (modo poly)

    @property
    def label(self):
        return f'{self.largura:.0f}×{self.altura:.0f}'


@dataclass
class PecaPos:
    """Peça posicionada numa chapa."""
    peca:    Peca
    x:       float
    y:       float
    largura: float   # pode ser altura original se girada
    altura:  float
    girada:  bool = False

    @property
    def x2(self): return self.x + self.largura
    @property
    def y2(self): return self.y + self.altura

    def to_dict(self):
        return {
            'id':      self.peca.id,
            'nome':    self.peca.nome,
            'x':       round(self.x, 2),
            'y':       round(self.y, 2),
            'w':       round(self.largura, 2),
            'h':       round(self.altura, 2),
            'girada':  self.girada,
            'cor':     self.peca.cor,
            'label':   f'{self.largura:.0f}×{self.altura:.0f}',
        }


@dataclass
class Chapa:
    indice:       int
    largura:      float
    altura:       float
    pecas:        List[PecaPos] = field(default_factory=list)

    @property
    def area_util(self):
        return sum(p.largura * p.altura for p in self.pecas)

    @property
    def aproveitamento(self):
        return self.area_util / (self.largura * self.altura)

    def to_dict(self):
        return {
            'indice':         self.indice,
            'largura':        self.largura,
            'altura':         self.altura,
            'pecas':          [p.to_dict() for p in self.pecas],
            'aproveitamento': round(self.aproveitamento * 100, 1),
            'n_pecas':        len(self.pecas),
        }


# ══════════════════════════════════════════════════════════════
#  ALGORITMO GUILLOTINE — para peças retangulares
# ══════════════════════════════════════════════════════════════

class Guillotine:
    """
    Algoritmo de corte guilhotina com heurística MAXRECTS-like.
    Mantém lista de espaços livres e divide após cada inserção.
    Ref: Jylänki (2010) "A Thousand Ways to Pack the Bin"
    """

    def __init__(self, largura, altura, margem_x=0, margem_y=0, gap=0):
        self.W  = largura
        self.H  = altura
        self.mx = margem_x
        self.my = margem_y
        self.gap = gap
        # Área útil
        ux = largura  - 2 * margem_x
        uy = altura   - 2 * margem_y
        self.livres = [{'x': margem_x, 'y': margem_y, 'w': ux, 'h': uy}]
        self.posicionadas: List[PecaPos] = []

    def _score(self, esp, pw, ph):
        """Heurística Best Short Side Fit (BSSF)."""
        leftover_h = esp['w'] - pw
        leftover_v = esp['h'] - ph
        short = min(leftover_h, leftover_v)
        long_ = max(leftover_h, leftover_v)
        return (short, long_)

    def inserir(self, peca: Peca) -> bool:
        """Tenta inserir a peça. Retorna True se conseguiu."""
        g = self.gap
        melhor_score = None
        melhor_esp   = None
        melhor_girada = False
        melhor_pw, melhor_ph = 0, 0

        orientacoes = [(peca.largura, peca.altura, False)]
        if peca.rotacao and peca.largura != peca.altura:
            orientacoes.append((peca.altura, peca.largura, True))

        for pw, ph, girada in orientacoes:
            pw_g = pw + g
            ph_g = ph + g
            for esp in self.livres:
                if pw_g <= esp['w'] + g and ph_g <= esp['h'] + g:
                    # cabe (com tolerância do gap)
                    if pw <= esp['w'] and ph <= esp['h']:
                        s = self._score(esp, pw, ph)
                        if melhor_score is None or s < melhor_score:
                            melhor_score  = s
                            melhor_esp    = esp
                            melhor_girada = girada
                            melhor_pw, melhor_ph = pw, ph

        if melhor_esp is None:
            return False

        # Posicionar
        e = melhor_esp
        pos = PecaPos(
            peca=peca, x=e['x'], y=e['y'],
            largura=melhor_pw, altura=melhor_ph,
            girada=melhor_girada,
        )
        self.posicionadas.append(pos)
        self._dividir(e, melhor_pw, melhor_ph)
        return True

    def _dividir(self, esp, pw, ph):
        """Divide o espaço livre em dois retângulos (corte guilhotina)."""
        self.livres.remove(esp)
        g = self.gap

        # Direita
        dw = esp['w'] - pw - g
        if dw > 0:
            self.livres.append({
                'x': esp['x'] + pw + g,
                'y': esp['y'],
                'w': dw,
                'h': esp['h'],
            })
        # Cima
        dh = esp['h'] - ph - g
        if dh > 0:
            self.livres.append({
                'x': esp['x'],
                'y': esp['y'] + ph + g,
                'w': pw,
                'h': dh,
            })

        self._merge_livres()

    def _merge_livres(self):
        """Remove espaços contidos em outros (simplificação)."""
        keep = []
        for i, a in enumerate(self.livres):
            dominado = False
            for j, b in enumerate(self.livres):
                if i != j:
                    if (b['x'] <= a['x'] and b['y'] <= a['y'] and
                            b['x']+b['w'] >= a['x']+a['w'] and
                            b['y']+b['h'] >= a['y']+a['h']):
                        dominado = True
                        break
            if not dominado:
                keep.append(a)
        self.livres = keep


# ══════════════════════════════════════════════════════════════
#  ALGORITMO BL-FILL — para polígonos arbitrários (DXF)
# ══════════════════════════════════════════════════════════════

def _bl_fill(pecas_input, sheet_w, sheet_h, gap, margin_x, margin_y):
    """Bottom-Left Fill com Shapely para geometrias arbitrárias."""
    from shapely.geometry import box
    from shapely import affinity

    def rotate_poly(poly, angle):
        r = affinity.rotate(poly, angle, origin=(0,0), use_radians=False)
        minx, miny = r.bounds[0], r.bounds[1]
        return affinity.translate(r, -minx, -miny)

    sheet_box = box(0, 0, sheet_w, sheet_h)
    placed    = []
    placed_polys = []

    def candidate_positions():
        pts = [(margin_x, margin_y)]
        for p in placed_polys:
            mx2, my2 = p.bounds[2], p.bounds[3]
            pts += [(mx2 + gap, margin_y), (margin_x, my2 + gap)]
        return sorted(set((round(x,1), round(y,1)) for x,y in pts),
                      key=lambda p: (p[1], p[0]))

    for pi in pecas_input:
        poly_orig = pi['polygon']
        best = None

        for angle in ([0, 90] if pi.get('rotacao', True) else [0]):
            poly_r = rotate_poly(poly_orig, angle)
            pw, ph = poly_r.bounds[2], poly_r.bounds[3]

            for cx, cy in candidate_positions():
                poly_t = affinity.translate(poly_r, cx, cy)
                if not sheet_box.contains(poly_t):
                    continue
                buffered = poly_t.buffer(gap/2)
                if any(buffered.intersects(p) for p in placed_polys):
                    continue
                score = (round(cy,1), round(cx,1))
                if best is None or score < best[0]:
                    best = (score, cx, cy, angle, poly_t, pw, ph)

        if best:
            _, bx, by, ba, bp, bw, bh = best
            placed.append({
                'nome':   pi['nome'],
                'x': bx, 'y': by,
                'w': bw, 'h': bh,
                'girada': ba == 90,
                'cor':    pi.get('cor','#4472C4'),
                'label':  f'{bw:.0f}×{bh:.0f}',
                'area_real': pi['area_real'],
            })
            placed_polys.append(bp)

    return placed


# ══════════════════════════════════════════════════════════════
#  FUNÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════

def run(pecas: List[Peca], sheet_w, sheet_h, gap, margin_x, margin_y,
        modo='rect') -> List[Chapa]:
    """
    Executa o nesting e retorna lista de Chapa.
    modo: 'rect' (guilhotina) | 'poly' (BL-fill Shapely)
    """
    # Atribuir cores
    nomes_unicos = list(dict.fromkeys(p.id for p in pecas))
    cor_map = {n: PALETA[i % len(PALETA)] for i, n in enumerate(nomes_unicos)}
    for p in pecas:
        p.cor = cor_map[p.id]

    # Expandir por quantidade, ordenar por área decrescente
    fila: List[Peca] = []
    for p in pecas:
        for _ in range(p.quantidade):
            fila.append(p)
    fila.sort(key=lambda p: p.largura * p.altura, reverse=True)

    chapas: List[Chapa] = []

    while fila:
        chapa = Chapa(indice=len(chapas)+1, largura=sheet_w, altura=sheet_h)
        nao_colocadas: List[Peca] = []

        if modo in ('rect', 'poly'):
            # Guillotine para ambos os modos — eficiente e correto
            # para bounding boxes retangulares (inclui DXFs).
            g = Guillotine(sheet_w, sheet_h, margin_x, margin_y, gap)
            for peca in fila:
                if not g.inserir(peca):
                    nao_colocadas.append(peca)
            chapa.pecas = g.posicionadas

        elif False:  # BL-Fill Shapely reservado para polígonos irregulares reais
            pi_list = [
                {'nome': p.nome, 'polygon': p.polygon,
                 'area_real': p.area_real, 'rotacao': p.rotacao,
                 'cor': p.cor}
                for p in fila
            ]
            colocadas_info = _bl_fill(
                pi_list, sheet_w, sheet_h, gap, margin_x, margin_y)
            colocados_nomes = [c['nome'] for c in colocadas_info]

            for p in fila:
                if p.nome in colocados_nomes:
                    colocados_nomes.remove(p.nome)
                    info = next(c for c in colocadas_info
                                if c['nome'] == p.nome)
                    chapa.pecas.append(PecaPos(
                        peca=p,
                        x=info['x'], y=info['y'],
                        largura=info['w'], altura=info['h'],
                        girada=info['girada'],
                    ))
                else:
                    nao_colocadas.append(p)

        if not chapa.pecas:
            break  # segurança: evitar loop infinito

        chapas.append(chapa)
        fila = nao_colocadas

    return chapas


def parse_dxf(filepath: str) -> dict:
    """Lê um DXF e retorna dict com polygon, área, bbox e nome."""
    import ezdxf
    from shapely.geometry import Polygon, LineString
    from shapely.ops import polygonize
    from shapely import affinity

    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()
    lines = []

    for e in msp:
        t = e.dxftype()
        if t == 'LINE':
            lines.append(LineString([
                (e.dxf.start.x, e.dxf.start.y),
                (e.dxf.end.x,   e.dxf.end.y),
            ]))
        elif t == 'LWPOLYLINE':
            pts = [(p[0], p[1]) for p in e.get_points()]
            if e.closed: pts.append(pts[0])
            for i in range(len(pts)-1):
                lines.append(LineString([pts[i], pts[i+1]]))
        elif t == 'ARC':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            sa, ea = math.radians(e.dxf.start_angle), math.radians(e.dxf.end_angle)
            if ea < sa: ea += 2*math.pi
            pts = [(cx + r*math.cos(sa + (ea-sa)*i/32),
                    cy + r*math.sin(sa + (ea-sa)*i/32)) for i in range(33)]
            for i in range(len(pts)-1):
                lines.append(LineString([pts[i], pts[i+1]]))
        elif t == 'CIRCLE':
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            pts = [(cx+r*math.cos(2*math.pi*i/64),
                    cy+r*math.sin(2*math.pi*i/64)) for i in range(65)]
            for i in range(len(pts)-1):
                lines.append(LineString([pts[i], pts[i+1]]))

    polys = list(polygonize(lines))
    if not polys:
        raise ValueError(f'Não foi possível reconstruir polígono: {filepath}')

    outer = max(polys, key=lambda p: p.area)
    holes_area = sum(p.area for p in polys if p != outer and outer.contains(p))
    area_real  = outer.area - holes_area

    minx, miny, maxx, maxy = outer.bounds
    norm = affinity.translate(outer, -minx, -miny)

    return {
        'polygon':  norm,
        'area_real': area_real,
        'width':    maxx - minx,
        'height':   maxy - miny,
        'nome':     os.path.splitext(os.path.basename(filepath))[0],
    }

def parse_dxf_bytes(filename: str, data: bytes) -> dict:
    """Lê um DXF a partir de bytes (upload em memória)."""
    import tempfile, os
    suffix = os.path.splitext(filename)[1] or '.dxf'
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(data)
        tmp_path = tmp.name
    try:
        result = parse_dxf(tmp_path)
        result['nome'] = os.path.splitext(filename)[0]
        return result
    finally:
        os.unlink(tmp_path)
