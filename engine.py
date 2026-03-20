"""
engine.py  —  Nesting 2D v3.0
Modos:
  rect : Guillotine (bounding box)
  poly : True-shape via No-Fit Polygon (Shapely)
         Rotação por peça: 0°, 90°, 180°, 270°, livre(360°)
         Furos e recortes respeitados
"""

import math, os
from dataclasses import dataclass, field
from typing import List

PALETA = [
    '#4472C4','#ED7D31','#70AD47','#FF0000','#7030A0',
    '#00B0F0','#FFC000','#FF00FF','#00B050','#C00000',
    '#5B9BD5','#92D050','#833C00','#203864','#548235',
]

# ── Dataclasses ───────────────────────────────────────────────────

@dataclass
class Peca:
    id:          str
    nome:        str
    largura:     float
    altura:      float
    area_real:   float
    quantidade:  int
    rotacao:     bool  = True
    angulo_max:  float = 360.0   # 0=fixo, 90, 180, 270, 360=livre
    cor:         str   = '#4472C4'
    polygon:     object = None   # Shapely Polygon exterior
    holes:       object = None   # list of Shapely Polygons


@dataclass
class PecaPos:
    peca:        Peca
    x:           float
    y:           float
    largura:     float
    altura:      float
    girada:      bool  = False
    angulo:      float = 0.0
    poly_coords: list  = field(default_factory=list)
    hole_coords: list  = field(default_factory=list)

    def to_dict(self):
        return {
            'id':          self.peca.id,
            'nome':        self.peca.nome,
            'x':           round(self.x, 2),
            'y':           round(self.y, 2),
            'w':           round(self.largura, 2),
            'h':           round(self.altura, 2),
            'girada':      self.girada,
            'angulo':      round(self.angulo, 1),
            'cor':         self.peca.cor,
            'label':       f'{self.largura:.0f}×{self.altura:.0f}',
            'poly_coords': self.poly_coords,
            'hole_coords': self.hole_coords,
        }


@dataclass
class Chapa:
    indice:  int
    largura: float
    altura:  float
    pecas:   List[PecaPos] = field(default_factory=list)

    @property
    def area_util(self):
        return sum(p.peca.area_real for p in self.pecas)

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


# ══════════════════════════════════════════════════════════════════
#  GUILLOTINE (modo rect)
# ══════════════════════════════════════════════════════════════════

class Guillotine:
    def __init__(self, w, h, mx, my, gap):
        ux, uy = w - 2*mx, h - 2*my
        self.livres = [{'x': mx, 'y': my, 'w': ux, 'h': uy}]
        self.gap = gap
        self.posicionadas: List[PecaPos] = []

    def inserir(self, peca: Peca) -> bool:
        orients = [(peca.largura, peca.altura, False)]
        if peca.rotacao and peca.largura != peca.altura:
            orients.append((peca.altura, peca.largura, True))
        best = None
        for pw, ph, gir in orients:
            for esp in self.livres:
                if pw <= esp['w'] and ph <= esp['h']:
                    s = (min(esp['w']-pw, esp['h']-ph),
                         max(esp['w']-pw, esp['h']-ph))
                    if best is None or s < best[0]:
                        best = (s, esp, pw, ph, gir)
        if best is None:
            return False
        _, esp, pw, ph, gir = best
        self.posicionadas.append(
            PecaPos(peca=peca, x=esp['x'], y=esp['y'],
                    largura=pw, altura=ph, girada=gir))
        self._dividir(esp, pw, ph)
        return True

    def _dividir(self, esp, pw, ph):
        self.livres.remove(esp)
        g = self.gap
        if esp['w'] - pw - g > 0:
            self.livres.append({'x': esp['x']+pw+g, 'y': esp['y'],
                                 'w': esp['w']-pw-g, 'h': esp['h']})
        if esp['h'] - ph - g > 0:
            self.livres.append({'x': esp['x'], 'y': esp['y']+ph+g,
                                 'w': pw, 'h': esp['h']-ph-g})
        keep = []
        for i, a in enumerate(self.livres):
            if not any(j != i and
                       b['x'] <= a['x'] and b['y'] <= a['y'] and
                       b['x']+b['w'] >= a['x']+a['w'] and
                       b['y']+b['h'] >= a['y']+a['h']
                       for j, b in enumerate(self.livres)):
                keep.append(a)
        self.livres = keep


# ══════════════════════════════════════════════════════════════════
#  TRUE-SHAPE NESTING (modo poly)
# ══════════════════════════════════════════════════════════════════

def _angles_for(angulo_max: float) -> list:
    if angulo_max <= 0:   return [0]
    if angulo_max <= 90:  return [0, 90]
    if angulo_max <= 180: return [0, 90, 180]
    if angulo_max <= 270: return [0, 90, 180, 270]
    return [0, 90, 180, 270]   # livre = 4 orientações principais


def _rotate_norm(poly, angle):
    from shapely import affinity
    r = affinity.rotate(poly, angle, origin=(0, 0), use_radians=False)
    b = r.bounds
    return affinity.translate(r, -b[0], -b[1])


def _candidates(placed, sheet_w, sheet_h, gap, mx, my):
    pts = [(mx, my)]
    for poly in placed:
        b = poly.bounds
        pts += [(b[2]+gap, my), (mx, b[3]+gap), (b[2]+gap, b[3]+gap)]
        for x, y in list(poly.exterior.coords)[::3]:
            pts += [(x+gap, y), (x, y+gap), (x+gap, y+gap)]
    seen, result = set(), []
    for cx, cy in sorted(pts, key=lambda p: (round(p[1],1), round(p[0],1))):
        k = (round(cx,1), round(cy,1))
        if k not in seen and cx >= mx and cy >= my:
            seen.add(k)
            result.append((cx, cy))
    return result


def _true_shape_pass(fila, sheet_w, sheet_h, gap, mx, my):
    from shapely.geometry import box as sbox
    from shapely import affinity

    sheet  = sbox(0, 0, sheet_w, sheet_h)
    placed = []   # polígonos buffered já posicionados
    items  = []   # PecaPos

    for peca in fila:
        poly_base = peca.polygon
        if poly_base is None:
            poly_base = sbox(0, 0, peca.largura, peca.altura)

        angles = _angles_for(peca.angulo_max)
        best   = None
        cands  = _candidates(placed, sheet_w, sheet_h, gap, mx, my)

        for angle in angles:
            poly_r = _rotate_norm(poly_base, angle)
            b = poly_r.bounds
            pw, ph = b[2]-b[0], b[3]-b[1]

            for cx, cy in cands:
                poly_t = affinity.translate(poly_r, cx, cy)

                if not sheet.contains(poly_t):
                    continue

                buf = poly_t.buffer(gap / 2.0, resolution=8)
                if any(buf.intersects(pp) and
                       buf.intersection(pp).area > 0.05
                       for pp in placed):
                    continue

                score = (round(cy, 1), round(cx, 1))
                if best is None or score < best[0]:
                    best = (score, cx, cy, angle, poly_t, pw, ph)

        if best is None:
            continue   # não coube — vai para próxima chapa

        _, bx, by, ba, bp, bw, bh = best

        # Coordenadas para SVG
        pc = [(round(x,2), round(y,2)) for x,y in bp.exterior.coords]
        hc = []
        if hasattr(bp, 'interiors'):
            for interior in bp.interiors:
                hc.append([(round(x,2), round(y,2)) for x,y in interior.coords])

        # Furos do DXF: transladar junto com a peça
        holes_placed = []
        if peca.holes:
            for h in peca.holes:
                h_r = _rotate_norm(h, ba)
                h_t = affinity.translate(h_r, bx, by)
                hc.append([(round(x,2), round(y,2)) for x,y in h_t.exterior.coords])
                holes_placed.append(h_t)

        items.append(PecaPos(
            peca=peca, x=bx, y=by, largura=bw, altura=bh,
            girada=(ba != 0), angulo=ba,
            poly_coords=pc, hole_coords=hc,
        ))
        placed.append(bp.buffer(gap/2.0, resolution=8))

    colocados_ids = {id(p.peca) for p in items}
    sobras = [p for p in fila if id(p) not in
              {id(item.peca) for item in items}]

    # Reconstruir sobras corretamente
    colocados_pecas = list(items)
    sobras = []
    colocados_count = {}
    for item in items:
        colocados_count[item.peca.id] = colocados_count.get(item.peca.id, 0) + 1

    fila_count = {}
    for p in fila:
        fila_count[p.id] = fila_count.get(p.id, 0) + 1

    for p in fila:
        restam = fila_count.get(p.id, 0) - colocados_count.get(p.id, 0)
        if restam > 0:
            sobras.append(p)
            colocados_count[p.id] = colocados_count.get(p.id, 0) + 1

    return items, sobras


# ══════════════════════════════════════════════════════════════════
#  FUNÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def run(pecas: List[Peca], sheet_w, sheet_h, gap, margin_x, margin_y,
        modo='rect') -> List[Chapa]:

    nomes = list(dict.fromkeys(p.id for p in pecas))
    cor_map = {n: PALETA[i % len(PALETA)] for i, n in enumerate(nomes)}
    for p in pecas:
        p.cor = cor_map[p.id]

    fila: List[Peca] = []
    for p in pecas:
        for _ in range(p.quantidade):
            fila.append(p)
    fila.sort(key=lambda p: p.area_real, reverse=True)

    tem_poly = any(p.polygon is not None for p in fila)
    usar_poly = (modo == 'poly') and tem_poly

    chapas: List[Chapa] = []
    max_chapas = 50   # limite de segurança

    while fila and len(chapas) < max_chapas:
        chapa = Chapa(indice=len(chapas)+1, largura=sheet_w, altura=sheet_h)

        if usar_poly:
            colocadas, fila = _true_shape_pass(
                fila, sheet_w, sheet_h, gap, margin_x, margin_y)
            chapa.pecas = colocadas
        else:
            g = Guillotine(sheet_w, sheet_h, margin_x, margin_y, gap)
            nao_col = []
            for peca in fila:
                if not g.inserir(peca):
                    nao_col.append(peca)
            chapa.pecas = g.posicionadas
            fila = nao_col

        if not chapa.pecas:
            break

        chapas.append(chapa)

    return chapas


# ══════════════════════════════════════════════════════════════════
#  PARSER DXF
# ══════════════════════════════════════════════════════════════════

def _arc_pts(cx, cy, r, a_start, a_end, n=32):
    s = math.radians(a_start)
    e = math.radians(a_end)
    if e < s:
        e += 2 * math.pi
    return [(cx + r*math.cos(s+(e-s)*i/n),
             cy + r*math.sin(s+(e-s)*i/n)) for i in range(n+1)]


def _collect_lines(msp):
    from shapely.geometry import LineString
    lines = []
    for e in msp:
        t = e.dxftype()
        try:
            if t == 'LINE':
                p1 = (e.dxf.start.x, e.dxf.start.y)
                p2 = (e.dxf.end.x,   e.dxf.end.y)
                if p1 != p2:
                    lines.append(LineString([p1, p2]))
            elif t == 'LWPOLYLINE':
                pts = [(p[0], p[1]) for p in e.get_points()]
                if e.closed and pts[0] != pts[-1]:
                    pts.append(pts[0])
                for i in range(len(pts)-1):
                    if pts[i] != pts[i+1]:
                        lines.append(LineString([pts[i], pts[i+1]]))
            elif t == 'POLYLINE':
                pts = [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
                if e.is_closed and pts[0] != pts[-1]:
                    pts.append(pts[0])
                for i in range(len(pts)-1):
                    if pts[i] != pts[i+1]:
                        lines.append(LineString([pts[i], pts[i+1]]))
            elif t == 'ARC':
                pts = _arc_pts(e.dxf.center.x, e.dxf.center.y,
                               e.dxf.radius, e.dxf.start_angle, e.dxf.end_angle)
                for i in range(len(pts)-1):
                    if pts[i] != pts[i+1]:
                        lines.append(LineString([pts[i], pts[i+1]]))
            elif t == 'CIRCLE':
                pts = _arc_pts(e.dxf.center.x, e.dxf.center.y,
                               e.dxf.radius, 0, 360, n=64)
                for i in range(len(pts)-1):
                    lines.append(LineString([pts[i], pts[i+1]]))
            elif t == 'SPLINE':
                try:
                    pts = [(p[0], p[1]) for p in e.flattening(0.1)]
                except Exception:
                    pts = [(p[0], p[1]) for p in e.control_points]
                for i in range(len(pts)-1):
                    if pts[i] != pts[i+1]:
                        lines.append(LineString([pts[i], pts[i+1]]))
            elif t == 'ELLIPSE':
                cx2, cy2 = e.dxf.center.x, e.dxf.center.y
                major = e.dxf.major_axis
                a = math.hypot(major.x, major.y)
                b2 = a * e.dxf.ratio
                ang = math.atan2(major.y, major.x)
                pts = []
                for i in range(65):
                    theta = 2*math.pi*i/64
                    x = cx2 + a*math.cos(theta)*math.cos(ang) - b2*math.sin(theta)*math.sin(ang)
                    y = cy2 + a*math.cos(theta)*math.sin(ang) + b2*math.sin(theta)*math.cos(ang)
                    pts.append((x, y))
                for i in range(len(pts)-1):
                    lines.append(LineString([pts[i], pts[i+1]]))
        except Exception:
            continue
    return lines


def parse_dxf(filepath: str) -> dict:
    import ezdxf
    from shapely.ops import polygonize, unary_union
    from shapely import affinity

    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()
    lines = _collect_lines(msp)

    if not lines:
        raise ValueError('Nenhuma geometria encontrada.')

    polys = list(polygonize(lines))
    if not polys:
        merged = unary_union(lines)
        polys  = list(polygonize(merged))
    if not polys:
        raise ValueError('Não foi possível reconstruir o polígono. '
                         'Verifique se o contorno está fechado.')

    polys.sort(key=lambda p: p.area, reverse=True)
    outer  = polys[0]
    holes  = [p for p in polys[1:] if p.area > 0.5 and outer.contains(p)]

    b = outer.bounds
    outer_n = affinity.translate(outer, -b[0], -b[1])
    holes_n = [affinity.translate(h, -b[0], -b[1]) for h in holes]

    area_real = outer_n.area - sum(h.area for h in holes_n)
    width, height = b[2]-b[0], b[3]-b[1]

    return {
        'polygon':     outer_n,
        'holes':       holes_n,
        'area_real':   area_real,
        'width':       width,
        'height':      height,
        'nome':        os.path.splitext(os.path.basename(filepath))[0],
        'poly_coords': [(round(x,2), round(y,2))
                        for x,y in outer_n.exterior.coords],
        'hole_coords': [[(round(x,2), round(y,2)) for x,y in h.exterior.coords]
                        for h in holes_n],
    }


def parse_dxf_bytes(filename: str, data: bytes) -> dict:
    import tempfile
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
