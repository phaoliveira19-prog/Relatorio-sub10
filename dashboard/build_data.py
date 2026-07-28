#!/usr/bin/env python3
"""Extracts data.json for the Sub-10 dashboard from the Sistema de gestao spreadsheet.

Usage: python3 build_data.py <caminho-para-o-xlsx> [saida.json]

Regras aplicadas (decisoes tomadas na conversa com o treinador):
- Presenca de um atleta num treino = ele tem nota preenchida em "Avaliacao
  individual do treino" naquele dia (blank = ausencia).
- Referencia de minutagem por sessao = 80 minutos (confirmado, nao 90).
- Conteudo principal / secundario / concorrente sao papeis pedagogicos
  distintos (todos intencionais) e sao mantidos separados, nunca somados
  num unico numero.
- Tendencia do atleta (ultimos 5 treinos vs media geral): limiar de +/-0.5.
"""
import json
import re
import sys
import unicodedata
from datetime import datetime, date
from pathlib import Path

import openpyxl

PHOTOS_DIR = Path(__file__).parent / 'photos'


def slug(name):
    s = unicodedata.normalize('NFKD', str(name)).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^a-zA-Z0-9]+', '-', s).strip('-').lower()
    return s

MONTH_PT = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
QUARTILE_ORDER = ['Q1', 'Q2', 'Q3', 'Q4']
TENURE_ORDER = ['0–3 meses', '4–6 meses', '7–12 meses', '13–24 meses', '+24 meses']
TASK_ORDER = ['Analítico', 'Situacional', 'Conceitual', 'Conceitual em ambiente específico', 'Jogo formal']
COMPLEX_ORDER = ['Baixa', 'Média', 'Alta']
PART_ORDER = ['Preparatória', 'Conceitual', 'Conexão']
TREND_THRESHOLD = 0.5
REFERENCE_MINUTES = 80


def d(dt):
    if dt is None:
        return None
    if isinstance(dt, datetime):
        dt = dt.date()
    return dt.isoformat()


def month_label(dt):
    if dt is None:
        return None
    if isinstance(dt, datetime):
        dt = dt.date()
    return f"{MONTH_PT[dt.month - 1]}/{str(dt.year)[2:]}"


def tenure_band(months):
    if months is None:
        return None
    if months <= 3:
        return '0–3 meses'
    if months <= 6:
        return '4–6 meses'
    if months <= 12:
        return '7–12 meses'
    if months <= 24:
        return '13–24 meses'
    return '+24 meses'


def months_between(start, end):
    if start is None:
        return None
    if isinstance(start, datetime):
        start = start.date()
    if isinstance(end, datetime):
        end = end.date()
    return (end.year - start.year) * 12 + (end.month - start.month) - (1 if end.day < start.day else 0)


def norm(s):
    return str(s).strip() if s is not None else ''


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    src = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 else 'data.json'

    wb = openpyxl.load_workbook(src, data_only=True)
    as_of = date.today().isoformat()

    # ---------- Banco de atletas ----------
    ws = wb['Banco de atletas']
    athletes = []
    athlete_by_nick = {}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if not row[0]:
            continue
        name, nickname, aid, dob, quartile, position, foot, joined, photo_link = (
            row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[9]
        )
        characteristic, improvement = row[10], row[11]
        tenure_months = months_between(joined, date.today()) if joined else None
        local_photo = next((PHOTOS_DIR / f"{slug(nickname)}{ext}" for ext in ('.jpg', '.jpeg', '.png')
                             if (PHOTOS_DIR / f"{slug(nickname)}{ext}").exists()), None)
        if local_photo:
            photo = f"./photos/{local_photo.name}"
        elif photo_link and str(photo_link).startswith('http') and 'drive.google.com' not in str(photo_link):
            photo = photo_link
        else:
            photo = None
        a = {
            'id': aid,
            'name': name,
            'nickname': nickname,
            'dob': d(dob),
            'birthMonth': month_label(dob) if dob else None,
            'quartile': quartile,
            'position': position,
            'foot': foot,
            'joined': d(joined),
            'tenureMonths': tenure_months,
            'tenureBand': tenure_band(tenure_months),
            'category': row[8],
            'photo': photo,
            'characteristic': characteristic,
            'improvement': improvement,
            'status': 'ativo',
        }
        athletes.append(a)
        athlete_by_nick[norm(nickname).lower()] = a

    # ---------- Banco de treinos ----------
    ws = wb['Banco de treinos']
    trainings = []
    blocks = []
    contents = []  # long format, role-tagged: principal / secundario / concorrente
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        (dt, session_id, category,
         c_prin_prep, c_sec_prep, c_conc_prep, tipo_prep, tarefa_prep, complex_prep,
         c_prin_conc, c_sec_conc, c_conc_conc, tipo_conc, tarefa_conc, complex_conc,
         relacao_conexao, tarefa_conexao, complex_conexao, _blank,
         nota_geral, num_atletas, atletas_avaliacao,
         min_prep, jog_ativos_prep, taa_prep, desemp_prep,
         min_conc, jog_ativos_conc, taa_conc, desemp_conc,
         min_conexao, jog_ativos_conexao, taa_conexao, desemp_conexao,
         observacoes, minutagem_efetiva, total_taa, pct_minutagem, pct_taa) = row
        if dt is None:
            continue  # sessoes sem data (pos-treino parcial, ainda nao detalhadas)
        month = month_label(dt)
        trainings.append({
            'date': d(dt), 'month': month, 'sessionId': session_id, 'category': category,
            'note': nota_geral, 'numAthletes': num_atletas, 'athletesEval': atletas_avaliacao,
            'effectiveMinutes': minutagem_efetiva, 'totalTAA': total_taa,
            'pctMinutagem': pct_minutagem if isinstance(pct_minutagem, (int, float)) else None,
            'pctTAA': pct_taa if isinstance(pct_taa, (int, float)) else None,
            'observations': observacoes,
        })

        parts = [
            ('Preparatória', c_prin_prep, c_sec_prep, c_conc_prep, tipo_prep, tarefa_prep, complex_prep,
             min_prep, jog_ativos_prep, taa_prep, desemp_prep, None),
            ('Conceitual', c_prin_conc, c_sec_conc, c_conc_conc, tipo_conc, tarefa_conc, complex_conc,
             min_conc, jog_ativos_conc, taa_conc, desemp_conc, None),
            ('Conexão', None, None, None, 'Conexão com jogo', tarefa_conexao, complex_conexao,
             min_conexao, jog_ativos_conexao, taa_conexao, desemp_conexao, relacao_conexao),
        ]
        for part, c_prin, c_sec, c_conc, tipo, tarefa, complexidade, minutes, jog_ativos, taa, desemp, relacao in parts:
            if minutes is None and c_prin is None and relacao is None:
                continue
            blocks.append({
                'date': d(dt), 'month': month, 'sessionId': session_id, 'part': part,
                'content': relacao if part == 'Conexão' else c_prin,
                'contentType': tipo, 'taskType': tarefa, 'complexity': complexidade,
                'minutes': minutes, 'numAthletesActive': jog_ativos, 'taa': taa,
                'performance': desemp, 'sessionNote': nota_geral,
            })
            for role, content in (('principal', c_prin), ('secundario', c_sec), ('concorrente', c_conc)):
                if content:
                    contents.append({
                        'date': d(dt), 'month': month, 'part': part, 'role': role,
                        'content': content, 'contentType': tipo,
                    })

    # ---------- Avaliacao individual do treino ----------
    ws = wb['Avaliação individual do treino']
    header = [c.value for c in ws[1]]
    date_cols = header[2:]
    notes = []
    training_by_date = {t['date']: t for t in trainings}
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        nickname = row[0]
        if not nickname:
            continue
        a = athlete_by_nick.get(norm(nickname).lower())
        for col_idx, raw_date in enumerate(date_cols, start=2):
            note = row[col_idx]
            if note is None:
                continue
            dt = datetime.strptime(raw_date, '%d/%m/%Y').date()
            notes.append({
                'athlete': nickname, 'date': dt.isoformat(), 'month': month_label(dt), 'note': note,
                'position': a['position'] if a else None,
                'quartile': a['quartile'] if a else None,
                'foot': a['foot'] if a else None,
                'tenureBand': a['tenureBand'] if a else None,
                'photo': a['photo'] if a else None,
            })

    # ---------- Pos-jogo ----------
    games = []
    if 'Pós-jogo' in wb.sheetnames:
        ws = wb['Pós-jogo']
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if row[0] is None:
                continue
            (game_no, dt, category, opponent, rank, opp_category, competition, local, city, period,
             fmt, platform, coach, captain, minutes, gf, ga, saldo, result, yellow, red, performance,
             comments, positives, negatives, gap_days) = row
            games.append({
                'gameNo': game_no, 'date': d(dt), 'month': month_label(dt), 'category': category,
                'opponent': opponent, 'rank': rank, 'oppCategory': opp_category, 'competition': competition,
                'local': local, 'city': city, 'period': period, 'format': fmt, 'platform': platform,
                'coach': coach, 'captain': captain, 'minutes': minutes, 'goalsFor': gf, 'goalsAgainst': ga,
                'saldo': saldo,
                'result': result if (gf is not None and ga is not None) else None,
                'yellow': yellow, 'red': red, 'performance': performance, 'comments': comments,
                'positives': positives, 'negatives': negatives, 'daysSinceLastGame': gap_days,
            })

    # ---------- Pos-jogo individual ----------
    game_ind = []
    if 'Pós-jogo individual' in wb.sheetnames:
        ws = wb['Pós-jogo individual']
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            if row[3] is None:
                continue
            (game_no, dt, category, name, position, titularity, minutes, performance,
             goals, goals_against, assists, yellow, red, comments) = row
            a = athlete_by_nick.get(norm(name).lower())
            game_ind.append({
                'gameNo': game_no, 'date': d(dt), 'month': month_label(dt), 'category': category,
                'athlete': name, 'position': position or (a['position'] if a else None),
                'quartile': a['quartile'] if a else None, 'foot': a['foot'] if a else None,
                'titularity': titularity, 'minutes': minutes or 0, 'performance': performance,
                'goals': goals or 0, 'goalsAgainst': goals_against or 0, 'assists': assists or 0,
                'yellow': yellow or 0, 'red': red or 0, 'comments': comments,
            })

    all_dates = [t['date'] for t in trainings] + [g['date'] for g in games]
    all_dated = trainings + games
    months_seen = sorted({t['month'] for t in all_dated if t['month']},
                          key=lambda m: min(t['date'] for t in all_dated if t['month'] == m))

    data = {
        'meta': {
            'generatedAt': datetime.now().isoformat(timespec='seconds'),
            'asOf': as_of,
            'minDate': min(all_dates) if all_dates else None,
            'maxDate': max(all_dates) if all_dates else None,
            'referenceMinutes': REFERENCE_MINUTES,
            'trendThreshold': TREND_THRESHOLD,
            'months': months_seen,
            'athletes': [a['nickname'] for a in athletes],
            'positions': sorted({a['position'] for a in athletes if a['position']}),
            'quartiles': QUARTILE_ORDER,
            'competitions': sorted({g['competition'] for g in games if g['competition']}),
            'ranks': sorted({g['rank'] for g in games if g['rank']}),
            'formats': sorted({g['format'] for g in games if g['format']}),
            'results': sorted({g['result'] for g in games if g['result']}),
            'taskTypes': sorted({bl['taskType'] for bl in blocks if bl['taskType']}),
            'complexities': sorted({bl['complexity'] for bl in blocks if bl['complexity']}),
            'parts': PART_ORDER,
            'contents': sorted({c['content'] for c in contents if c['content']}),
            'quartileOrder': QUARTILE_ORDER,
            'tenureOrder': TENURE_ORDER,
            'taskOrder': TASK_ORDER,
            'complexityOrder': COMPLEX_ORDER,
            'partOrder': PART_ORDER,
        },
        'athletes': athletes,
        'trainings': trainings,
        'blocks': blocks,
        'contents': contents,
        'notes': notes,
        'games': games,
        'gameInd': game_ind,
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=None, separators=(',', ':'))

    print(f"OK: {out_path}")
    print(f"  athletes={len(athletes)} trainings={len(trainings)} blocks={len(blocks)} "
          f"contents={len(contents)} notes={len(notes)} games={len(games)} gameInd={len(game_ind)}")


if __name__ == '__main__':
    main()
