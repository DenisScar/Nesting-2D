"""
app.py  —  Nesting 2D v3.0
"""

import os, io, threading, webbrowser, tempfile
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
from engine import run, parse_dxf_bytes, Peca
import export

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

IS_PROD = os.environ.get('RENDER', False)
OUTPUT  = Path(tempfile.gettempdir()) / 'nesting_output'
OUTPUT.mkdir(parents=True, exist_ok=True)

_job = {'status': 'idle', 'progresso': '', 'resultado': None,
        'erro': None, 'chapas': [], 'resumo': {}}

# DXFs da sessão: {nome: {'bytes': bytes, 'info': dict}}
_dxf_store = {}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload_dxf', methods=['POST'])
def upload_dxf():
    arquivos = request.files.getlist('dxfs')
    if not arquivos:
        return jsonify({'erro': 'Nenhum arquivo recebido.'}), 400

    pecas, erros = [], []
    for f in arquivos:
        nome = os.path.basename(f.filename)
        if not nome.lower().endswith('.dxf'):
            erros.append({'arquivo': nome, 'erro': 'Não é um arquivo .dxf'})
            continue
        try:
            data = f.read()
            info = parse_dxf_bytes(nome, data)
            _dxf_store[nome] = {'bytes': data, 'info': info}
            pecas.append({
                'nome':        info['nome'],
                'arquivo':     nome,
                'largura':     round(info['width'],    1),
                'altura':      round(info['height'],   1),
                'area':        round(info['area_real'], 1),
                'poly_coords': info.get('poly_coords', []),
                'hole_coords': info.get('hole_coords', []),
            })
        except Exception as e:
            erros.append({'arquivo': nome, 'erro': str(e)})

    return jsonify({'pecas': pecas, 'erros': erros})


@app.route('/api/remover_dxf', methods=['POST'])
def remover_dxf():
    nome = (request.json or {}).get('nome', '')
    _dxf_store.pop(nome, None)
    return jsonify({'ok': True})


@app.route('/api/rodar', methods=['POST'])
def rodar():
    global _job
    if _job['status'] == 'rodando':
        return jsonify({'erro': 'Cálculo em andamento.'}), 400
    _job = {'status': 'rodando', 'progresso': 'Iniciando...',
            'resultado': None, 'erro': None, 'chapas': [], 'resumo': {}}
    t = threading.Thread(target=_executar, args=(request.json,), daemon=True)
    t.start()
    return jsonify({'ok': True})


@app.route('/api/status')
def status():
    return jsonify({k: _job[k] for k in
                    ['status', 'progresso', 'resultado', 'erro', 'resumo']})


@app.route('/api/chapas')
def get_chapas():
    return jsonify({'chapas': _job['chapas']})


@app.route('/api/baixar')
def baixar():
    p = _job.get('resultado')
    if not p or not os.path.exists(p):
        return 'Arquivo não encontrado', 404
    return send_file(p, as_attachment=True,
                     download_name=os.path.basename(p))


def _executar(dados):
    global _job
    try:
        nome    = dados.get('nome_projeto', 'Projeto').strip() or 'Projeto'
        modo    = dados.get('modo', 'poly')
        sheet_w = float(dados['chapa_x'])
        sheet_h = float(dados['chapa_y'])
        gap     = float(dados['gap'])
        mx      = float(dados['margem_x'])
        my      = float(dados['margem_y'])

        config = {
            'nome': nome, 'sheet_w': sheet_w, 'sheet_h': sheet_h,
            'gap': gap, 'margin_x': mx, 'margin_y': my,
            'material':  dados.get('material', ''),
            'espessura': float(dados.get('espessura', 1.5)),
            'rho':       float(dados.get('rho', 7850)),
            'preco_kg':  float(dados.get('preco_kg', 0)),
        }

        peca_cfg = {p['nome']: p for p in dados.get('pecas', [])}
        pecas, pecas_export = [], []

        if modo == 'poly':
            if not _dxf_store:
                raise ValueError('Nenhum DXF carregado.')

            for i, (arq_nome, entry) in enumerate(_dxf_store.items()):
                _job['progresso'] = f'Preparando peça {i+1}/{len(_dxf_store)}: {arq_nome}'
                info = entry['info']
                cfg  = peca_cfg.get(info['nome'], {})
                qtd  = int(cfg.get('quantidade', 1))
                ang  = float(cfg.get('angulo_max', 360))

                pecas.append(Peca(
                    id=info['nome'], nome=info['nome'],
                    largura=info['width'], altura=info['height'],
                    area_real=info['area_real'],
                    quantidade=qtd, angulo_max=ang,
                    polygon=info['polygon'],
                    holes=info.get('holes', []),
                ))
                pecas_export.append({
                    'nome':       info['nome'],
                    'largura':    round(info['width'], 1),
                    'altura':     round(info['height'], 1),
                    'quantidade': qtd, 'cor': '',
                })
        else:
            for pc in dados.get('pecas', []):
                pecas.append(Peca(
                    id=pc['nome'], nome=pc['nome'],
                    largura=float(pc['largura']), altura=float(pc['altura']),
                    area_real=float(pc['largura']) * float(pc['altura']),
                    quantidade=int(pc['quantidade']),
                    rotacao=bool(pc.get('rotacao', True)),
                ))
                pecas_export.append({
                    'nome':       pc['nome'],
                    'largura':    float(pc['largura']),
                    'altura':     float(pc['altura']),
                    'quantidade': int(pc['quantidade']), 'cor': '',
                })

        if not pecas:
            raise ValueError('Nenhuma peça configurada.')

        n_total = sum(p.quantidade for p in pecas)
        _job['progresso'] = f'Calculando nesting ({n_total} peças)...'
        chapas = run(pecas, sheet_w, sheet_h, gap, mx, my, modo=modo)

        _job['progresso'] = f'Gerando relatório ({len(chapas)} chapas)...'
        cor_map = {p.id: p.cor for p in pecas}
        for pc in pecas_export:
            pc['cor'] = cor_map.get(pc['nome'], '#4472C4')

        chapas_dict = [c.to_dict() for c in chapas]
        _job['chapas'] = chapas_dict

        out_file = OUTPUT / f'nesting_{nome.replace(" ", "_")}.xlsx'
        export.gerar(chapas_dict, pecas_export, config, str(out_file))

        n_aloc = sum(len(c.pecas) for c in chapas)
        aprov  = (sum(c.aproveitamento for c in chapas) / len(chapas) * 100
                  if chapas else 0)

        _job.update({
            'status':    'concluido',
            'progresso': 'Concluído!',
            'resultado': str(out_file),
            'resumo': {
                'n_chapas': len(chapas),
                'n_total':  n_total,
                'n_aloc':   n_aloc,
                'aprov':    f'{aprov:.1f}%',
                'arquivo':  out_file.name,
            },
        })

    except Exception as e:
        import traceback; traceback.print_exc()
        _job.update({'status': 'erro', 'erro': str(e),
                     'progresso': f'Erro: {e}'})


if __name__ == '__main__':
    def _open():
        import time; time.sleep(1.2)
        webbrowser.open('http://localhost:5000')
    threading.Thread(target=_open, daemon=True).start()
    print('\n' + '─'*50)
    print('  Nesting 2D v3.0  →  http://localhost:5000')
    print('─'*50 + '\n')
    app.run(debug=False, port=5000, use_reloader=False)
