"""
export.py  —  Nesting 2D v3.5
Correções:
  - Título e nome do projeto separados visualmente (sem sobreposição)
  - Espessura adicionada aos parâmetros
  - "Cortes" renomeado para "Rec. internos"
  - Tabela de peças por chapa REMOVIDA (estava duplicada)
  - Sobra reaproveitável mantida em todas as chapas quando detectável
"""

import io, math
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import PathPatch
from matplotlib.path import Path as MPath

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image as RLImage, PageBreak, HRFlowable,
)

# ── Cores ─────────────────────────────────────────────────────────
BLU_D  = colors.HexColor('#1F497D')
BLU_M  = colors.HexColor('#4472C4')
BLU_L  = colors.HexColor('#DCE6F1')
GRN_L  = colors.HexColor('#E2EFDA')
GOLD   = colors.HexColor('#7F6000')
GOLD_L = colors.HexColor('#FFF2CC')
GRAY_M = colors.HexColor('#CCCCCC')
WHITE  = colors.white


def _rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16)/255 for i in (0, 2, 4))


def _darker(h, f=0.6):
    r, g, b = _rgb(h)
    return (r*f, g*f, b*f)


# ── Comprimento de corte ──────────────────────────────────────────

def _poly_len(coords):
    t = 0.0
    for i in range(len(coords)-1):
        t += math.hypot(coords[i+1][0]-coords[i][0], coords[i+1][1]-coords[i][1])
    return t


def _cut_mm(p):
    t = 0.0
    pc = p.get('poly_coords', [])
    if pc:
        t += _poly_len(pc)
    for h in p.get('hole_coords', []):
        t += _poly_len(h)
    return t or 2*(p.get('w', 0)+p.get('h', 0))


def _n_internos(p):
    """Número de recortes internos (furos). Não conta o contorno externo."""
    return len(p.get('hole_coords', []))


# ── Remnant ───────────────────────────────────────────────────────

def _remnant(sw, sh, pecas, mx, my):
    """
    Identifica a maior sobra retangular na chapa.
    Retorna (w, h) ou None. Presente sempre que sobra > 20mm.
    """
    if not pecas:
        return None
    max_x = max(p['x']+p['w'] for p in pecas)
    rw = sw - max_x - mx
    if rw > 20:
        return (round(rw), round(sh))
    max_y = max(p['y']+p['h'] for p in pecas)
    rh = sh - max_y - my
    if rh > 20:
        return (round(sw), round(rh))
    return None


# ── Render matplotlib ─────────────────────────────────────────────

def _render(chapa, sw, sh, mx, my, idx, total):
    fw, fh = 8.0, 8.0*(sh/sw)
    fig, ax = plt.subplots(figsize=(fw, fh))
    ax.set_facecolor('#E8EDF4')
    fig.patch.set_facecolor('#FFFFFF')

    ax.add_patch(mpatches.Rectangle((0,0), sw, sh,
        lw=2, ec='#1F497D', fc='#F5F5F5', zorder=1))
    ax.add_patch(mpatches.Rectangle((0,0), sw, sh,
        lw=0, fc='none', hatch='///', ec='#DDDDDD', zorder=2, alpha=0.4))
    ax.add_patch(mpatches.Rectangle((mx,my), sw-2*mx, sh-2*my,
        lw=0.8, ec='#AAAAAA', fc='none', ls='--', zorder=3))

    for p in chapa['pecas']:
        cor = p['cor']
        brd = _darker(cor)
        pc  = p.get('poly_coords', [])
        hc  = p.get('hole_coords', [])

        if pc and len(pc) >= 3:
            verts = [(x,y) for x,y in pc]
            codes = [MPath.MOVETO]+[MPath.LINETO]*(len(verts)-2)+[MPath.CLOSEPOLY]
            for hole in hc:
                if len(hole) >= 3:
                    verts += [(x,y) for x,y in hole]
                    codes += [MPath.MOVETO]+[MPath.LINETO]*(len(hole)-2)+[MPath.CLOSEPOLY]
            ax.add_patch(PathPatch(MPath(verts, codes),
                lw=0.9, ec=brd, fc=(*_rgb(cor), 0.82), zorder=4))
            xs = [v[0] for v in pc]; ys = [v[1] for v in pc]
            cx = (min(xs)+max(xs))/2; cy = (min(ys)+max(ys))/2
        else:
            ax.add_patch(mpatches.Rectangle((p['x'],p['y']), p['w'], p['h'],
                lw=0.9, ec=brd, fc=(*_rgb(cor), 0.82), zorder=4))
            cx = p['x']+p['w']/2; cy = p['y']+p['h']/2

        fs = max(4, min(8, p['w']/55, p['h']/18))
        ax.text(cx, cy, p['nome'], ha='center', va='center',
                fontsize=fs, fontweight='bold', color='white', zorder=5,
                bbox=dict(boxstyle='round,pad=0.1', fc='none', ec='none'))

    ax.annotate('', xy=(sw, -sh*.025), xytext=(0, -sh*.025),
                arrowprops=dict(arrowstyle='<->', color='#555', lw=0.8))
    ax.text(sw/2, -sh*.04, f'{sw:.0f} mm', ha='center', va='top', fontsize=6.5, color='#555')
    ax.annotate('', xy=(sw*1.025, sh), xytext=(sw*1.025, 0),
                arrowprops=dict(arrowstyle='<->', color='#555', lw=0.8))
    ax.text(sw*1.03, sh/2, f'{sh:.0f} mm', ha='left', va='center',
            fontsize=6.5, color='#555', rotation=90)

    ax.set_xlim(-sw*.02, sw*1.10); ax.set_ylim(-sh*.07, sh*1.04)
    ax.set_aspect('equal'); ax.axis('off')
    ax.set_title(f'Chapa {idx}/{total}  —  {chapa["n_pecas"]} peças  |  '
                 f'Aproveitamento: {chapa["aproveitamento"]:.1f}%',
                 fontsize=9, fontweight='bold', color='#1F497D', pad=5)
    plt.tight_layout(pad=0.2)
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf


# ── Estilos ───────────────────────────────────────────────────────

def _S():
    def ps(name, **kw):
        return ParagraphStyle(name, **kw)
    return {
        # Capa — título e nome bem separados
        'h1':   ps('h1',  fontSize=20, textColor=BLU_D, fontName='Helvetica-Bold',
                   alignment=1, spaceAfter=0),
        'proj': ps('proj',fontSize=13, textColor=BLU_M, fontName='Helvetica-Bold',
                   alignment=1, spaceBefore=6, spaceAfter=0),
        'dt':   ps('dt',  fontSize=8,  textColor=colors.gray,
                   fontName='Helvetica', alignment=1, spaceBefore=4),
        # Seções
        'sec':  ps('sec', fontSize=9,  textColor=WHITE,
                   fontName='Helvetica-Bold', leftIndent=6, leading=13),
        # Corpo
        'n':    ps('n',   fontSize=8.5, fontName='Helvetica', leading=12),
        'sm':   ps('sm',  fontSize=7.5, fontName='Helvetica', leading=10,
                   textColor=colors.HexColor('#333333')),
        'b':    ps('b',   fontSize=8.5, fontName='Helvetica-Bold', leading=12),
        # KPIs
        'kv':   ps('kv',  fontSize=15, fontName='Helvetica-Bold',
                   textColor=BLU_D, alignment=1),
        'kl':   ps('kl',  fontSize=7,  fontName='Helvetica',
                   textColor=colors.HexColor('#5A6A8A'), alignment=1),
    }


def _shdr(txt, s, tw):
    t = Table([[Paragraph(txt, s['sec'])]], colWidths=[tw])
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), BLU_D),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ]))
    return t


def _kpis(items, s, tw):
    cw = tw/len(items)
    t = Table(
        [[Paragraph(str(v), s['kv']) for v,l in items],
         [Paragraph(l,      s['kl']) for v,l in items]],
        colWidths=[cw]*len(items))
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), BLU_L),
        ('BOX',        (0,0), (-1,-1), 0.5, BLU_M),
        ('INNERGRID',  (0,0), (-1,-1), 0.3, BLU_M),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('TOPPADDING',    (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    return t


# ── Geração principal ─────────────────────────────────────────────

def gerar(chapas_dict, pecas_cfg, config, output_path):
    pw, ph = A4
    mg = 15*mm
    tw = pw - 2*mg

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=mg, rightMargin=mg,
        topMargin=mg, bottomMargin=18*mm,
        title=f"Nesting — {config.get('nome','')}",
    )
    s  = _S()
    st = []

    sw    = config.get('sheet_w', 1200)
    sh    = config.get('sheet_h', 2000)
    esp   = config.get('espessura', 1.5)
    rho   = config.get('rho', 7850)
    pkg   = config.get('preco_kg', 0)
    mx    = config.get('margin_x', 40)
    my    = config.get('margin_y', 90)
    nc    = len(chapas_dict)
    aprov = (sum(c['aproveitamento'] for c in chapas_dict)/nc if nc else 0)
    massa = (sw/1000)*(sh/1000)*(esp/1000)*rho*nc
    custo = massa*pkg
    nome_p   = config.get('nome', '')
    datahora = datetime.now().strftime('%d/%m/%Y  %H:%M')

    # ── CAPA ─────────────────────────────────────────────────────
    # Título e nome bem separados — linha horizontal entre eles
    st += [
        Spacer(1, 10*mm),
        Paragraph('RELATÓRIO DE NESTING', s['h1']),
        Spacer(1, 3*mm),
        HRFlowable(width='60%', thickness=1, color=BLU_M, hAlign='CENTER'),
        Spacer(1, 3*mm),
        Paragraph(nome_p, s['proj']),
        Paragraph(datahora, s['dt']),
        Spacer(1, 6*mm),
        HRFlowable(width='100%', thickness=1.5, color=BLU_M),
        Spacer(1, 5*mm),
    ]

    # KPIs
    kl = [
        (nc,               'Chapas'),
        (f'{aprov:.1f}%',  'Aproveitamento médio'),
        (f'{massa:.1f} kg','Peso total (chapas)'),
    ]
    if pkg:
        kl.append((f'R$ {custo:.2f}', 'Custo MP'))
    st.append(_kpis(kl, s, tw))
    st.append(Spacer(1, 5*mm))

    # ── PARÂMETROS ───────────────────────────────────────────────
    st.append(_shdr('PARÂMETROS', s, tw))
    st.append(Spacer(1, 1*mm))
    pd_ = [
        ['Material',         config.get('material','—'),
         'Chapa X × Y (mm)', f'{sw:.0f} × {sh:.0f}'],
        ['Espessura (mm)',    f'{esp}',          # ← corrigido
         'Gap (mm)',          str(config.get('gap', 0))],
        ['Margem X / Y (mm)', f'{mx} / {my}',
         'Densidade (kg/m³)', str(rho)],
        ['Preço MP (R$/kg)',  f'{pkg:.2f}' if pkg else '—',
         '', ''],
    ]
    cw_p = [tw*f for f in (.20, .30, .20, .30)]
    pt = Table(pd_, colWidths=cw_p)
    pt.setStyle(TableStyle([
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,0), (-1,-1), 8),
        ('FONTNAME',      (0,0), (0,-1),  'Helvetica-Bold'),
        ('FONTNAME',      (2,0), (2,-1),  'Helvetica-Bold'),
        ('BACKGROUND',    (0,0), (0,-1),  BLU_L),
        ('BACKGROUND',    (2,0), (2,-1),  BLU_L),
        ('GRID',          (0,0), (-1,-1), 0.3, GRAY_M),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING',   (0,0), (-1,-1), 6),
    ]))
    st += [pt, Spacer(1, 5*mm)]

    # ── LISTA DE PEÇAS ────────────────────────────────────────────
    st.append(_shdr('LISTA DE PEÇAS', s, tw))
    st.append(Spacer(1, 1*mm))

    # Agregar de todas as chapas
    pi = {}
    for chapa in chapas_dict:
        for p in chapa['pecas']:
            nm = p['nome']
            if nm not in pi:
                pi[nm] = {
                    'w': p['w'], 'h': p['h'],
                    'area': p['w']*p['h'],
                    'cut1': _cut_mm(p),
                    'nint': _n_internos(p),   # rec. internos
                    'qtd':  0,
                    'cor':  p['cor'],
                }
            pi[nm]['qtd'] += 1

    # Ordenar por área decrescente
    pord = sorted(pi.items(), key=lambda x: x[1]['area'], reverse=True)

    hdr = [Paragraph(t, s['b']) for t in [
        '#', 'Peça', 'L × A (mm)', 'Área (m²)',
        'Rec. internos',            # ← renomeado
        'Comp./unid.', 'Comp. total', 'Qtd.',
    ]]
    rows = [hdr]
    tot_cut = 0.0

    for i, (nm, info) in enumerate(pord):
        cu = info['cut1'] / 1000
        ct = cu * info['qtd']
        tot_cut += ct
        rows.append([
            Paragraph(str(i+1),                              s['sm']),
            Paragraph(nm,                                    s['n']),
            Paragraph(f"{info['w']:.1f} × {info['h']:.1f}", s['sm']),
            Paragraph(f"{info['w']*info['h']/1e6:.4f}",     s['sm']),
            Paragraph(str(info['nint']),                     s['sm']),
            Paragraph(f"{cu:.3f} m",                         s['sm']),
            Paragraph(f"{ct:.3f} m",                         s['sm']),
            Paragraph(str(info['qtd']),                      s['sm']),
        ])

    # Linha de totais
    rows.append([
        Paragraph('',                                              s['b']),
        Paragraph('TOTAL',                                         s['b']),
        Paragraph('',                                              s['b']),
        Paragraph('',                                              s['b']),
        Paragraph('',                                              s['b']),
        Paragraph('',                                              s['b']),
        Paragraph(f"{tot_cut:.3f} m",                             s['b']),
        Paragraph(str(sum(i['qtd'] for _, i in pord)),            s['b']),
    ])

    cw_l = [tw*f for f in (.04, .24, .14, .10, .10, .13, .13, .07)]
    lt = Table(rows, colWidths=cw_l, repeatRows=1)
    lt.setStyle(TableStyle([
        ('FONTSIZE',       (0,0), (-1,-1), 7.5),
        ('GRID',           (0,0), (-1,-1), 0.3, GRAY_M),
        ('BACKGROUND',     (0,0), (-1, 0), BLU_D),
        ('TEXTCOLOR',      (0,0), (-1, 0), WHITE),
        ('BACKGROUND',     (0,-1),(-1,-1), GRN_L),
        ('FONTNAME',       (0,-1),(-1,-1), 'Helvetica-Bold'),
        ('ALIGN',          (2,0), (-1,-1), 'CENTER'),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [BLU_L, WHITE]),
        ('TOPPADDING',     (0,0), (-1,-1), 3),
        ('BOTTOMPADDING',  (0,0), (-1,-1), 3),
        ('LEFTPADDING',    (0,0), (-1,-1), 4),
    ]))
    st += [lt, Spacer(1, 2*mm)]

    # Caixa de resumo de corte
    rc = Table([[
        Paragraph(f'Comprimento total de corte: <b>{tot_cut:.3f} m</b>', s['n']),
        Paragraph(f'Tipos de peça: <b>{len(pord)}</b>',                  s['n']),
        Paragraph(f'Total de peças: <b>{sum(i["qtd"] for _,i in pord)}</b>', s['n']),
    ]], colWidths=[tw/3]*3)
    rc.setStyle(TableStyle([
        ('BACKGROUND',    (0,0), (-1,-1), GOLD_L),
        ('BOX',           (0,0), (-1,-1), 0.5,  GOLD),
        ('INNERGRID',     (0,0), (-1,-1), 0.3,  GOLD),
        ('TOPPADDING',    (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING',   (0,0), (-1,-1), 8),
    ]))
    st.append(rc)

    # ── UMA PÁGINA POR CHAPA ──────────────────────────────────────
    for chapa in chapas_dict:
        st.append(PageBreak())
        idx = chapa['indice']

        st.append(_shdr(
            f'CHAPA {idx} DE {nc}  —  {chapa["n_pecas"]} peças  |  '
            f'Aproveitamento: {chapa["aproveitamento"]:.1f}%', s, tw))
        st.append(Spacer(1, 3*mm))

        # Layout visual
        buf = _render(chapa, sw, sh, mx, my, idx, nc)
        ih  = min((ph - 4*mg) * 0.65, tw * (sh/sw))
        iw  = ih * (sw/sh)
        if iw > tw:
            iw = tw; ih = iw * (sh/sw)
        st += [RLImage(buf, width=iw, height=ih), Spacer(1, 3*mm)]

        # Sobra reaproveitável — sempre que detectada
        rem = _remnant(sw, sh, chapa['pecas'], mx, my)
        if rem:
            ra = rem[0]*rem[1]/1e6
            rb = Table([[Paragraph(
                f'Sobra reaproveitável estimada: '
                f'<b>{rem[0]:.0f} × {rem[1]:.0f} mm</b>'
                f'  |  Área: <b>{ra:.3f} m²</b>', s['n'])]],
                colWidths=[tw])
            rb.setStyle(TableStyle([
                ('BACKGROUND',    (0,0), (-1,-1), GOLD_L),
                ('BOX',           (0,0), (-1,-1), 0.5, GOLD),
                ('TOPPADDING',    (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('LEFTPADDING',   (0,0), (-1,-1), 8),
            ]))
            st += [rb, Spacer(1, 2*mm)]

        # ← Tabela de peças por chapa REMOVIDA (estava duplicada)

    # Rodapé
    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 7)
        canvas.setFillColor(colors.gray)
        canvas.drawString(mg, 10*mm,
                          f'Nesting 2D  —  {nome_p}  —  {datahora}')
        canvas.drawRightString(pw - mg, 10*mm, f'Página {doc.page}')
        canvas.restoreState()

    doc.build(st, onFirstPage=_footer, onLaterPages=_footer)
    return output_path
