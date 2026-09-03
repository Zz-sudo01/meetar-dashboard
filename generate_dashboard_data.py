#!/usr/bin/env python3
import json
import os
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
BASE_FILE = Path('/Users/zz/Desktop/Meetar底数.xlsx')
SPEND_FILE = Path('/Users/zz/Desktop/Meetar日耗.xlsx')
OUT_FILE = ROOT / 'dashboard-data.json'
OUT_JS_FILE = ROOT / 'dashboard-data.js'
OUT_ENC_FILE = ROOT / 'dashboard-data.enc'
EXCLUDED_MEDIA = {'自然量', '卸载重装'}
ENCRYPTION_ITERATIONS = '200000'

NS = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def parse_shared_strings(zf):
    if 'xl/sharedStrings.xml' not in zf.namelist():
        return []
    root = ET.fromstring(zf.read('xl/sharedStrings.xml'))
    return [''.join(t.text or '' for t in si.iterfind('.//a:t', NS)) for si in root.findall('a:si', NS)]


def sheet_target(zf, sheet_name=None, index=0):
    wb = ET.fromstring(zf.read('xl/workbook.xml'))
    rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    relmap = {rel.attrib['Id']: rel.attrib['Target'].lstrip('/') for rel in rels}
    sheets = wb.find('a:sheets', NS)
    sheet = None
    if sheet_name is None:
        sheet = sheets[index]
    else:
        for s in sheets:
            if s.attrib['name'] == sheet_name:
                sheet = s
                break
    if sheet is None:
        raise ValueError(f'sheet not found: {sheet_name or index}')
    target = relmap[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]
    if not target.startswith('xl/'):
        target = 'xl/' + target
    return target


def cell_text(cell, sst):
    t = cell.attrib.get('t')
    v = cell.find('a:v', NS)
    if t == 's' and v is not None:
        return sst[int(v.text)]
    if t == 'inlineStr':
        return ''.join(tn.text or '' for tn in cell.iterfind('.//a:t', NS))
    return v.text if v is not None else None


def read_rows(path, sheet_name=None, index=0):
    with zipfile.ZipFile(path) as zf:
        sst = parse_shared_strings(zf)
        target = sheet_target(zf, sheet_name=sheet_name, index=index)
        root = ET.fromstring(zf.read(target))
        rows = []
        for row in root.find('a:sheetData', NS).findall('a:row', NS):
            vals = []
            for cell in row.findall('a:c', NS):
                vals.append(cell_text(cell, sst))
            rows.append(vals)
        return rows


def iso_from_yyyymmdd(value):
    if not value:
        return None
    s = str(value).strip()
    if len(s) != 8 or not s.isdigit():
        return s
    return f'{s[:4]}-{s[4:6]}-{s[6:8]}'


def to_float(value):
    if value in (None, '', '\\N'):
        return None
    try:
        return float(value)
    except Exception:
        return None


def to_int(value):
    if value in (None, '', '\\N'):
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def clean_text(value):
    if value is None:
        return None
    s = str(value).strip()
    return s or None


def fill_missing_spend_campaigns(items):
    campaign_names = defaultdict(set)
    first_named_date = {}
    for row in items:
        name = clean_text(row.get('广告系列名称'))
        if name:
            key = (row['国家'], row['媒体类型'])
            campaign_names[key].add(name)
            row_date = row.get('时间')
            if row_date:
                first_named_date[key] = min(first_named_date.get(key, row_date), row_date)

    for row in items:
        if clean_text(row.get('广告系列名称')):
            continue
        key = (row['国家'], row['媒体类型'])
        names = sorted(campaign_names.get(key, set()))
        row_date = row.get('时间')
        named_date = first_named_date.get(key)
        if names and named_date and row_date and row_date >= named_date:
            row['广告系列名称'] = names[0] if len(names) == 1 else '未填写广告系列'
    return items


def load_base():
    rows = read_rows(BASE_FILE, index=0)
    header = rows[0]
    items = []
    for row in rows[1:]:
        record = {header[i]: row[i] if i < len(row) else None for i in range(len(header))}
        media_type = clean_text(record.get('媒体类型'))
        if media_type in EXCLUDED_MEDIA:
            continue
        out = {
            '时间': iso_from_yyyymmdd(record.get('日期')),
            'apm': clean_text(record.get('apm')),
            '媒体类型': media_type,
            '国家': clean_text(record.get('国家')),
            '广告系列名称': clean_text(record.get('广告系列名称')),
            '归因设备数': to_float(record.get('归因设备数')) or 0,
            '注册设备数': to_float(record.get('注册设备数')) or 0,
            '进入首页设备数': to_float(record.get('进入首页设备数')) or 0,
            '注册人数': to_float(record.get('注册人数')) or 0,
            '进入首页人数': to_float(record.get('进入首页人数')) or 0,
            '首日充值人数': to_float(record.get('首日充值人数')) or 0,
            '首日充值次数': to_float(record.get('首日充值次数')) or 0,
            '首日充值金额（美元）': to_float(record.get('首日充值金额（美元）')) or 0,
            '首日arpu（美元）': to_float(record.get('首日arpu（美元）')),
            '首日arppu（美元）': to_float(record.get('首日arppu（美元）')),
            '累计充值人数': to_float(record.get('累计充值人数')) or 0,
            '累计充值次数': to_float(record.get('累计充值次数')) or 0,
            '累计汇总充值金额（美元）': to_float(record.get('累计汇总充值金额（美元）')) or 0,
            '累计充值金额（美元）': to_float(record.get('累计充值金额（美元）')) or 0,
            '累计买币金额（美元）': to_float(record.get('累计买币金额（美元）')) or 0,
            '累计arpu（美元）': to_float(record.get('累计arpu（美元）')),
            '累计arppu（美元）': to_float(record.get('累计arppu（美元）')),
        }
        for col in ['ltv0', 'ltv1', 'ltv3', 'ltv7', 'ltv14', 'ltv30', 'ltv60', 'ltv90', 'ltv120',
                    'arpu0', 'arpu1', 'arpu3', 'arpu7', 'arpu14', 'arpu30', 'arpu60', 'arpu90', 'arpu120']:
            out[col] = to_float(record.get(col))
        items.append(out)
    return items


def load_spend():
    rows = read_rows(SPEND_FILE, sheet_name='Sheet1')
    if not rows:
        return []
    header = rows[0]
    channel_map = {
        'TT': 'tiktok',
        'gg': 'google',
        'FB': 'facebook',
        'SC': 'snapchat',
    }
    items = []
    for row in rows[1:]:
        record = {header[i]: row[i] if i < len(row) else None for i in range(len(header))}
        channel = clean_text(record.get('渠道')) or '未知'
        items.append({
            '时间': iso_from_yyyymmdd(record.get('日期')),
            '国家': clean_text(record.get('国家')) or '未知',
            '媒体类型': channel_map.get(channel, channel),
            '广告系列名称': clean_text(record.get('广告系列')) or '',
            '费用': to_float(record.get('消耗')) or 0,
        })
    return fill_missing_spend_campaigns(items)


def aggregate(items, keys):
    groups = {}
    for row in items:
        key = tuple(row.get(k) for k in keys)
        if key not in groups:
            groups[key] = {k: row.get(k) for k in keys}
            groups[key]['费用'] = 0.0
            groups[key]['count'] = 0
            groups[key]['_rows'] = []
        groups[key]['费用'] += row.get('费用', 0) or 0
        groups[key]['count'] += 1
        groups[key]['_rows'].append(row)
    return list(groups.values())


def add_derived(row, ltv_cols):
    spend = row.get('费用') or 0
    devices = row.get('进入首页设备数') or 0
    payers = row.get('首日充值人数') or 0
    row['进入首页设备成本'] = spend / devices if devices else None
    row['首日充值成本'] = spend / payers if payers else None
    row['首日充值率'] = payers / devices if devices else None
    row['首日ROI'] = (row.get('ltv0') or 0) / spend if spend and row.get('ltv0') is not None else None
    for col in ltv_cols[1:]:
        row[f'ROI{col[3:]}'] = (row.get(col) or 0) / spend if spend and row.get(col) is not None else None
    for prev, curr in zip(ltv_cols, ltv_cols[1:]):
        a = row.get(curr)
        b = row.get(prev)
        row[f'{curr.upper()}/{prev.upper()}'] = a / b if a is not None and b not in (None, 0) else None
    return row


def main():
    base = load_base()
    spend = load_spend()
    base_dates = sorted({r['时间'] for r in base if r['时间']})
    spend_dates = sorted({r['时间'] for r in spend if r['时间']})
    all_dates = sorted(set(base_dates) | set(spend_dates))
    common_dates = sorted(set(base_dates) & set(spend_dates))
    latest_day = common_dates[-1] if common_dates else (all_dates[-1] if all_dates else None)

    spend_by_date = defaultdict(float)
    spend_by_group = defaultdict(float)
    spend_by_series = defaultdict(float)
    spend_group_rows = defaultdict(list)
    for row in spend:
        spend_by_date[row['时间']] += row['费用']
        spend_by_group[(row['时间'], row['国家'], row['媒体类型'])] += row['费用']
        spend_by_series[(row['时间'], row['国家'], row['媒体类型'], row['广告系列名称'])] += row['费用']
        spend_group_rows[(row['时间'], row['国家'], row['媒体类型'])].append(row)

    # Group spend series for ranking views.
    series_groups = defaultdict(float)
    series_meta = {}
    for row in spend:
        key = (row['广告系列名称'], row['国家'], row['媒体类型'])
        series_groups[key] += row['费用']
        series_meta[key] = {'广告系列名称': key[0], '国家': key[1], '媒体类型': key[2]}

    spend_daily_groups = defaultdict(float)
    for row in spend:
        spend_daily_groups[(row['时间'], row['国家'], row['媒体类型'])] += row['费用']

    base_group_counts = defaultdict(int)
    for row in base:
        base_group_counts[(row['时间'], row['国家'], row['媒体类型'])] += 1

    # Base rows with daily spend attached by date + country + media type.
    detail = []
    for row in base:
        out = dict(row)
        key = (row['时间'], row['国家'], row['媒体类型'])
        count = base_group_counts[key] or 1
        out['费用'] = (spend_by_group[key] / count) if spend_by_group[key] else 0
        detail.append(add_derived(out, ['ltv0', 'ltv1', 'ltv3', 'ltv7', 'ltv14', 'ltv30', 'ltv60', 'ltv90', 'ltv120']))

    # Preserve spend that has no matching base row by creating a spend-only synthetic row.
    for key, spend_rows in spend_group_rows.items():
        if base_group_counts.get(key, 0):
            continue
        synthetic = {
            '时间': key[0],
            'apm': 'MEETAR',
            '媒体类型': key[2],
            '国家': key[1],
            '广告系列名称': clean_text(spend_rows[0].get('广告系列名称')) or '',
            '归因设备数': 0,
            '注册设备数': 0,
            '进入首页设备数': 0,
            '注册人数': 0,
            '进入首页人数': 0,
            '首日充值人数': 0,
            '首日充值次数': 0,
            '首日充值金额（美元）': 0,
            '首日arpu（美元）': None,
            '首日arppu（美元）': None,
            '累计充值人数': 0,
            '累计充值次数': 0,
            '累计汇总充值金额（美元）': 0,
            '累计充值金额（美元）': 0,
            '累计买币金额（美元）': 0,
            '累计arpu（美元）': None,
            '累计arppu（美元）': None,
        }
        for col in ['ltv0', 'ltv1', 'ltv3', 'ltv7', 'ltv14', 'ltv30', 'ltv60', 'ltv90', 'ltv120',
                    'arpu0', 'arpu1', 'arpu3', 'arpu7', 'arpu14', 'arpu30', 'arpu60', 'arpu90', 'arpu120']:
            synthetic[col] = None
        synthetic['费用'] = spend_by_group[key]
        detail.append(add_derived(synthetic, ['ltv0', 'ltv1', 'ltv3', 'ltv7', 'ltv14', 'ltv30', 'ltv60', 'ltv90', 'ltv120']))

    daily = []
    for date in all_dates:
        rows = [r for r in detail if r['时间'] == date]
        if not rows:
            continue
        agg = {
            '时间': date,
            '广告系列名称': '分日大汇总',
            'apm': 'MEETAR',
            '媒体类型': '全部',
            '国家': '全部',
            '费用': sum(r['费用'] or 0 for r in rows),
            '进入首页设备数': sum(r['进入首页设备数'] or 0 for r in rows),
            '首日充值人数': sum(r['首日充值人数'] or 0 for r in rows),
            'ltv0': sum(r['ltv0'] or 0 for r in rows if r['ltv0'] is not None) if any(r['ltv0'] is not None for r in rows) else None,
            'ltv1': sum(r['ltv1'] or 0 for r in rows if r['ltv1'] is not None) if any(r['ltv1'] is not None for r in rows) else None,
            'ltv3': sum(r['ltv3'] or 0 for r in rows if r['ltv3'] is not None) if any(r['ltv3'] is not None for r in rows) else None,
            'ltv7': sum(r['ltv7'] or 0 for r in rows if r['ltv7'] is not None) if any(r['ltv7'] is not None for r in rows) else None,
            'ltv14': sum(r['ltv14'] or 0 for r in rows if r['ltv14'] is not None) if any(r['ltv14'] is not None for r in rows) else None,
            'ltv30': sum(r['ltv30'] or 0 for r in rows if r['ltv30'] is not None) if any(r['ltv30'] is not None for r in rows) else None,
            'ltv60': sum(r['ltv60'] or 0 for r in rows if r['ltv60'] is not None) if any(r['ltv60'] is not None for r in rows) else None,
            'ltv90': sum(r['ltv90'] or 0 for r in rows if r['ltv90'] is not None) if any(r['ltv90'] is not None for r in rows) else None,
            'ltv120': sum(r['ltv120'] or 0 for r in rows if r['ltv120'] is not None) if any(r['ltv120'] is not None for r in rows) else None,
        }
        daily.append(add_derived(agg, ['ltv0', 'ltv1', 'ltv3', 'ltv7', 'ltv14', 'ltv30', 'ltv60', 'ltv90', 'ltv120']))

    match_check = [
        {'检查项': '底数行数', '数值': len(base)},
        {'检查项': '消耗行数', '数值': len(spend)},
        {'检查项': '底数日期数', '数值': len(base_dates)},
        {'检查项': '消耗日期数', '数值': len(spend_dates)},
        {'检查项': '最新共同日期', '数值': latest_day},
        {'检查项': '底数最新日期', '数值': base_dates[-1] if base_dates else None},
        {'检查项': '消耗最新日期', '数值': spend_dates[-1] if spend_dates else None},
    ]

    meta = {
        'latest_day': latest_day,
        'base_latest_day': base_dates[-1] if base_dates else None,
        'spend_latest_day': spend_dates[-1] if spend_dates else None,
        'base_source': str(BASE_FILE),
        'spend_source': str(SPEND_FILE),
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'note': 'Meetar 私有仪表盘：底数来自桌面 Meetar底数.xlsx，消耗来自桌面 Meetar日耗.xlsx。已过滤自然量和卸载重装，消耗按国家+渠道挂到底数。',
    }

    payload = {
        'meta': meta,
        'daily': daily,
        'detail': detail,
        'spendDaily': [
            {'时间': k[0], '国家': k[1], '媒体类型': k[2], '费用': v}
            for k, v in sorted(spend_daily_groups.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][2]))
        ],
        'spendSeries': [{'广告系列名称': k[0], '国家': k[1], '媒体类型': k[2], '费用': v} for k, v in sorted(series_groups.items(), key=lambda kv: kv[1], reverse=True)],
        'matchCheck': match_check,
    }

    serialized = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    OUT_FILE.write_text(serialized, encoding='utf-8')
    passphrase = os.environ.get('MEETAR_DASHBOARD_PASSWORD')
    if passphrase:
        subprocess.run([
            'openssl', 'enc', '-aes-256-cbc', '-pbkdf2', '-iter', ENCRYPTION_ITERATIONS,
            '-salt', '-base64', '-in', str(OUT_FILE), '-out', str(OUT_ENC_FILE),
            '-pass', f'pass:{passphrase}'
        ], check=True)
    else:
        print('warning: MEETAR_DASHBOARD_PASSWORD not set; skipped encrypted output')
    print(f'wrote {OUT_FILE}')
    if passphrase:
        print(f'wrote {OUT_ENC_FILE}')


if __name__ == '__main__':
    main()
