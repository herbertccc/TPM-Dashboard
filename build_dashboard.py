#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# TPM Dashboard Build Script
# Reads data from DingTalk API, computes metrics, injects into index.html
import csv
import json
import os
import re
import ssl
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta

try:
    import openpyxl
except ImportError:
    openpyxl = None

DATA_DIR = 'data'
DINGTALK_APP_KEY = os.environ.get('DINGTALK_APP_KEY', '')
DINGTALK_APP_SECRET = os.environ.get('DINGTALK_APP_SECRET', '')
DINGTALK_NODE_ID = 'P0MALyR8kNnB9yZliY70z99KJ3bzYmDO'
TREND_DAYS = 30

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

CLOSED = '\u5df2\u5173\u95ed'
RESOLVED = '\u5df2\u89e3\u51b3'
REOPEN = '\u91cd\u65b0\u6253\u5f00'
CLOSED_SET = (CLOSED, RESOLVED)
HIGH_RISK = '\u9ad8\u98ce\u9669\u95ee\u9898'
OVER90 = '\u8d8590\u5929'
UNKNOWN = '\u672a\u77e5'
IN_PROGRESS = '\u8fdb\u884c\u4e2d'
GENERAL = '\u901a\u7528'


def dingtalk_get_token():
    url = ('https://oapi.dingtalk.com/gettoken?appkey='
           + DINGTALK_APP_KEY + '&appsecret=' + DINGTALK_APP_SECRET)
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        if data.get('errcode') == 0:
            print('token ok')
            return data['access_token']
        raise Exception('token failed: ' + str(data))


def dingtalk_get_doc_info(access_token):
    # Try multiple API endpoints
    endpoints = [
        'https://api.dingtalk.com/v1.0/doc/workspaces/nodes/' + DINGTALK_NODE_ID,
        'https://api.dingtalk.com/v2.0/storage/spaces/files/' + DINGTALK_NODE_ID,
    ]
    headers = {'x-acs-dingtalk-access-token': access_token}
    for url in endpoints:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                print('doc info from: ' + url)
                return data
        except Exception as e:
            print('endpoint failed: ' + url.split('/')[-2] + '/' + url.split('/')[-1] + ' -> ' + str(e))
    return {}


def dingtalk_download_file(access_token, doc_info):
    doc_id = (doc_info.get('docId') or doc_info.get('id')
              or doc_info.get('dentryUuid') or DINGTALK_NODE_ID)
    # Try export API
    endpoints = [
        ('POST', 'https://api.dingtalk.com/v1.0/doc/documents/' + str(doc_id) + '/export',
         json.dumps({'targetFormat': 'xlsx'}).encode('utf-8')),
        ('POST', 'https://api.dingtalk.com/v1.0/doc/suites/documents/' + str(doc_id) + '/export',
         json.dumps({'targetFormat': 'xlsx'}).encode('utf-8')),
    ]
    headers = {
        'x-acs-dingtalk-access-token': access_token,
        'Content-Type': 'application/json'
    }
    for method, url, body in endpoints:
        try:
            req = urllib.request.Request(url, data=body, headers=headers, method=method)
            with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                file_url = (data.get('result', {}).get('fileUrl')
                            or data.get('fileUrl'))
                if file_url:
                    req2 = urllib.request.Request(file_url)
                    with urllib.request.urlopen(req2, timeout=60, context=SSL_CTX) as r2:
                        file_data = r2.read()
                        os.makedirs(DATA_DIR, exist_ok=True)
                        fp = os.path.join(DATA_DIR, 'dingtalk_data.xlsx')
                        with open(fp, 'wb') as f:
                            f.write(file_data)
                        print('downloaded: ' + fp + ' (' + str(len(file_data)) + ' bytes)')
                        return fp
                print('no url in response: ' + json.dumps(data, ensure_ascii=False)[:200])
        except Exception as e:
            print('export failed: ' + str(e))
    return None


def fetch_dingtalk_data():
    if not DINGTALK_APP_KEY or not DINGTALK_APP_SECRET:
        print('no dingtalk credentials, skip')
        return []
    print('--- fetch via DingTalk API ---')
    try:
        token = dingtalk_get_token()
        doc_info = dingtalk_get_doc_info(token)
        if not doc_info:
            print('doc info empty, skip')
            return []
        print('doc: ' + json.dumps(doc_info, ensure_ascii=False)[:300])
        fp = dingtalk_download_file(token, doc_info)
        if fp:
            return read_xlsx_rows(fp)
    except Exception as e:
        print('dingtalk API error: ' + str(e))
    print('will try local data instead')
    return []


def read_xlsx_rows(path):
    if not openpyxl:
        print('openpyxl not installed')
        return []
    rows = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        headers = []
        for c in ws[1]:
            headers.append(str(c.value).strip() if c.value else '')
        col_idx = {}
        for i, h in enumerate(headers):
            if h:
                col_idx[h] = i
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in row):
                continue
            d = {}
            for name, idx in col_idx.items():
                val = row[idx] if idx < len(row) else None
                d[name] = str(val).strip() if val is not None else ''
            rows.append(d)
        wb.close()
        print(os.path.basename(path) + ': ' + str(len(rows)) + ' rows')
    except Exception as e:
        print('read error ' + os.path.basename(path) + ': ' + str(e))
    return rows


def read_local_data():
    all_rows = []
    if not os.path.exists(DATA_DIR):
        return all_rows
    for fn in sorted(os.listdir(DATA_DIR)):
        fp = os.path.join(DATA_DIR, fn)
        if fn.endswith('.csv'):
            try:
                with open(fp, encoding='utf-8-sig') as f:
                    first = f.readline()
                    if not first or '\x00' in first:
                        continue
                    f.seek(0)
                    for row in csv.DictReader(f):
                        d = {}
                        for k, v in row.items():
                            d[k.strip() if k else ''] = (v.strip() if v else '')
                        all_rows.append(d)
                print(fn + ': ' + str(len(all_rows)) + ' rows')
            except Exception as e:
                print(fn + ' error: ' + str(e))
        elif fn.endswith('.xlsx') and openpyxl:
            all_rows.extend(read_xlsx_rows(fp))
    return all_rows


def parse_date(s):
    if not s:
        return None
    s = s.strip()
    if not s:
        return None
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S',
                '%Y/%m/%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    try:
        serial = float(s)
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        return dt.strftime('%Y-%m-%d')
    except (ValueError, OverflowError):
        pass
    return None


def parse_float(s, default=0.0):
    if not s:
        return default
    try:
        return float(s)
    except ValueError:
        return default


def is_expired(b):
    if b['status'] != CLOSED:
        return False
    if b.get('over90') != OVER90:
        return False
    cd = parse_date(b.get('created'))
    if cd:
        days = (datetime.now() - datetime.strptime(cd, '%Y-%m-%d')).days
        if days > 90:
            return True
    return False


def process_rows(rows):
    bugs = []
    col_proj = None
    col_val = None
    for key in rows[0].keys() if rows else []:
        if key and 'DI' in key and 'project' in key.lower():
            col_proj = key
        if key and 'DI' in key and key.endswith('\u503c'):
            col_val = key
    for r in rows:
        project = ''
        for k in ('DI-\u9879\u76ee', 'DB-\u9879\u76ee'):
            v = r.get(k, '').strip()
            if v:
                project = v
                break
        if not project:
            project = GENERAL
        di_val = parse_float(r.get('DB-DI\u503c', '0'))
        status = r.get('DB-\u4efb\u52a1\u72b6\u6001', '').strip()
        created = parse_date(r.get('\u521b\u5efa\u65f6\u95f4', ''))
        di_date = parse_date(r.get('DI-Date', ''))
        sla_val = r.get('DB-SLA\u8d85\u65f6', '').strip()
        bugs.append({
            'id': r.get('\u4efb\u52a1ID', '').strip(),
            'title': r.get('\u6807\u9898', '').strip(),
            'level': r.get('BUG\u7b49\u7ea7', '').strip() or 'none',
            'status': status,
            'assignee': r.get('\u6267\u884c\u8005', '').strip(),
            'created': created or '',
            'di_date': di_date or '',
            'di_value': di_val,
            'project': project,
            'owner': r.get('DB-Owner', '').strip(),
            'dept': r.get('DB-\u90e8\u95e8', '').strip(),
            'role': r.get('DB-\u89d2\u8272', '').strip(),
            'sla_timeout': 1 if sla_val else 0,
            'sla_days': parse_float(r.get('DB-SLA\u5929\u6570', '0')),
            'over90': r.get('\u89e3\u51b3\u8d8590\u5929', '').strip(),
            'risk_ctrl': r.get('\u91ca\u653e\u7ba1\u63a7', '').strip(),
        })
    return bugs


def compute_projects_data(bugs):
    by_project = defaultdict(list)
    for b in bugs:
        if is_expired(b):
            continue
        by_project[b['project']].append(b)
    projects_data = {}
    for proj in sorted(by_project.keys()):
        pb = by_project[proj]
        open_di = round(sum(b['di_value'] for b in pb
                            if b['status'] not in CLOSED_SET), 2)
        resolved_cnt = sum(1 for b in pb if b['status'] in CLOSED_SET)
        total = len(pb)
        sla = sum(1 for b in pb if b['sla_timeout'] > 0)
        block = sum(1 for b in pb if b['status'] not in CLOSED_SET
                    and b['risk_ctrl'] == HIGH_RISK)
        rr = '{:.1f}'.format(resolved_cnt / total * 100) if total > 0 else '0.0'
        preg = sum(1 for b in pb if b['status'] == REOPEN)
        if open_di < 20 and sla <= 2:
            health = 'normal'
        elif open_di < 40 and sla <= 5:
            health = 'warning'
        elif open_di < 80 and sla <= 20:
            health = 'caution'
        else:
            health = 'danger'
        js_bugs = []
        for b in pb:
            bod = b['di_value'] if b['status'] not in CLOSED_SET else 0.0
            js_bugs.append({
                'id': b['id'], 'title': b['title'], 'level': b['level'],
                'status': b['status'], 'assignee': b['assignee'],
                'created': b['created'], 'openDI': round(bod, 2),
                'sla_timeout': b['sla_timeout'], 'sla_days': b['sla_days'],
                'db_dept': b['dept'], 'db_role': b['role'],
            })
        projects_data[proj] = {
            'bugs': js_bugs,
            'stats': {
                'health': health, 'open_di': open_di,
                'resolve_rate': rr, 'pending_regression': preg,
                'sla': sla, 'block': block,
                'resolved': resolved_cnt, 'total': total,
            }
        }
    return projects_data


def compute_trends(bugs, projects_data):
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    dates = [(today - timedelta(days=i)).strftime('%Y-%m-%d')
             for i in range(TREND_DAYS - 1, -1, -1)]
    labels = [d[5:] for d in dates]
    valid = [b for b in bugs if not is_expired(b)]
    dnew = []
    dres = []
    for d in dates:
        dnew.append(round(sum(b['di_value'] for b in valid
                             if b['created'] == d), 2))
        dres.append(round(sum(b['di_value'] for b in valid
                              if b['status'] in CLOSED_SET
                              and b['di_date'] == d), 2))
    cum = []
    for i, d in enumerate(dates):
        if i == 0:
            cum.append(round(sum(b['di_value'] for b in valid
                                 if b['created'] <= d
                                 and b['status'] not in CLOSED_SET), 2))
        else:
            cum.append(round(cum[i - 1] + dnew[i] - dres[i], 2))
    pt = {}
    for proj in sorted(projects_data.keys()):
        pb = [b for b in valid if b['project'] == proj]
        pn = []
        pr = []
        for d in dates:
            pn.append(round(sum(b['di_value'] for b in pb
                                if b['created'] == d), 2))
            pr.append(round(sum(b['di_value'] for b in pb
                                if b['status'] in CLOSED_SET
                                and b['di_date'] == d), 2))
        pc = []
        for i, d in enumerate(dates):
            if i == 0:
                pc.append(round(sum(b['di_value'] for b in pb
                                    if b['created'] <= d
                                    and b['status'] not in CLOSED_SET), 2))
            else:
                pc.append(round(pc[i - 1] + pn[i] - pr[i], 2))
        pt[proj] = [pc, pn, pr]
    return {'dates': labels, 'cumulative': cum, 'daily_new': dnew,
            'daily_resolved': dres, 'project_trends': pt}


def compute_person_stats(bugs, projects_data):
    valid = [b for b in bugs if not is_expired(b)]
    pm = {}
    for b in valid:
        owner = b['owner'] or b['assignee']
        if not owner:
            continue
        if owner not in pm:
            pm[owner] = {
                'name': owner, 'role': b['role'] or UNKNOWN,
                'proj_contrib': {},
                'total_open_di': 0, 'total_solved_di': 0,
                'total_unsolved_di': 0, 'total_sla': 0,
                'total_reopen': 0, 'solved_tickets': 0,
                'total_tickets': 0, 'sla_days_sum': 0,
                'sla_days_count': 0,
            }
        p = pm[owner]
        p['total_tickets'] += 1
        if b['status'] in CLOSED_SET:
            p['solved_tickets'] += 1
            p['total_solved_di'] += b['di_value']
        elif b['status'] == REOPEN:
            p['total_reopen'] += 1
            p['total_unsolved_di'] += b['di_value']
        else:
            p['total_open_di'] += b['di_value']
            p['total_unsolved_di'] += b['di_value']
        p['total_sla'] += b['sla_timeout']
        if b['sla_days'] > 0:
            p['sla_days_sum'] += b['sla_days']
            p['sla_days_count'] += 1
        proj = b['project']
        if proj not in p['proj_contrib']:
            p['proj_contrib'][proj] = {'solved_di': 0, 'open_di': 0,
                                       'unsolved_di': 0}
        if b['status'] in CLOSED_SET:
            p['proj_contrib'][proj]['solved_di'] += b['di_value']
        elif b['status'] == REOPEN:
            p['proj_contrib'][proj]['unsolved_di'] += b['di_value']
        else:
            p['proj_contrib'][proj]['open_di'] += b['di_value']
            p['proj_contrib'][proj]['unsolved_di'] += b['di_value']
    for p in pm.values():
        p['total_open_di'] = round(p['total_open_di'], 2)
        p['total_solved_di'] = round(p['total_solved_di'], 2)
        p['total_unsolved_di'] = round(p['total_unsolved_di'], 2)
        for proj in p['proj_contrib']:
            for k in p['proj_contrib'][proj]:
                p['proj_contrib'][proj][k] = round(
                    p['proj_contrib'][proj][k], 2)
    return list(pm.values())


def read_milestones():
    path = os.path.join('milestone', '\u91cc\u7a0b\u7891\u8282\u70b9.xlsx')
    if not os.path.exists(path) or not openpyxl:
        return {}
    ms = {}
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            items = []
            for row in rows[1:]:
                if not any(v is not None for v in row):
                    continue
                dv = row[0] if len(row) > 0 else None
                nv = row[1] if len(row) > 1 else None
                sv = row[2] if len(row) > 2 else ''
                rv = row[3] if len(row) > 3 else ''
                if nv is None:
                    continue
                ds = ''
                if dv:
                    try:
                        serial = float(dv)
                        ds = (datetime(1899, 12, 30)
                              + timedelta(days=serial)).strftime('%Y-%m-%d')
                    except (ValueError, TypeError):
                        ds = str(dv)
                items.append({
                    'date': ds, 'name': str(nv).strip(),
                    'item': str(rv).strip() if rv else '',
                    'status': str(sv).strip() if sv else IN_PROGRESS,
                })
            ms[sn] = items
        wb.close()
        print('milestones: ' + str(sum(len(v) for v in ms.values())))
    except Exception as e:
        print('milestone error: ' + str(e))
    return ms


def read_version_plan():
    path = os.path.join('Version_Plan', 'Version_Plan.xlsx')
    if not os.path.exists(path) or not openpyxl:
        return []
    plans = []
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2:
                continue
            for row in rows[1:]:
                if not any(v is not None for v in row):
                    continue
                plans.append({
                    'date': str(row[0]).strip() if row[0] else '',
                    'name': sn,
                    'overview': str(row[1]).strip() if len(row) > 1 and row[1] else '',
                    'link': str(row[2]).strip() if len(row) > 2 and row[2] else '',
                })
        wb.close()
        print('version plans: ' + str(len(plans)))
    except Exception as e:
        print('version plan error: ' + str(e))
    return plans


def inject_data(html_path, data):
    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()
    reps = {
        'TREND_DATES': data['trend_dates'],
        'CUMULATIVE_DI': data['cumulative_di'],
        'DAILY_NEW': data['daily_new'],
        'DAILY_RESOLVED': data['daily_resolved'],
        'PROJECTS_DATA': data['projects_data'],
        'PROJECT_NAMES': data['project_names'],
        'PROJECT_TRENDS': data['project_trends'],
        'MILESTONES': data['milestones'],
        'PERSON_STATS': data['person_stats'],
        'VERSION_PLAN': data['version_plan'],
    }
    for var_name, value in reps.items():
        js = 'const ' + var_name + ' = ' + json.dumps(value, ensure_ascii=False) + ';'
        pat = r'const\s+' + var_name + r'\s*=\s*[\s\S]*?;'
        if re.search(pat, html):
            html = re.sub(pat, js, html)
        else:
            html = html.replace('<script>\n', '<script>\n' + js + '\n', 1)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)


def main():
    print('=' * 50)
    print('TPM Dashboard Build')
    print('=' * 50)
    rows = fetch_dingtalk_data()
    if not rows:
        print('error: no data from DingTalk API')
        return
    print('total rows: ' + str(len(rows)))
    bugs = process_rows(rows)
    print('valid bugs: ' + str(len(bugs)))
    projects_data = compute_projects_data(bugs)
    print('projects: ' + str(len(projects_data)))
    trends = compute_trends(bugs, projects_data)
    person_stats = compute_person_stats(bugs, projects_data)
    milestones = read_milestones()
    version_plan = read_version_plan()
    project_names = sorted(projects_data.keys())
    for proj in project_names:
        s = projects_data[proj]['stats']
        print(proj + ': OPEN DI=' + str(s['open_di'])
              + ', SLA=' + str(s['sla'])
              + ', BLOCK=' + str(s['block'])
              + ', rate=' + str(s['resolve_rate'])
              + '%, health=' + s['health'])
    print('staff: ' + str(len(person_stats)))
    print('trend days: ' + str(len(trends['dates'])))
    html_path = 'index.html'
    if not os.path.exists(html_path):
        print('error: ' + html_path + ' not found')
        return
    inject_data(html_path, {
        'trend_dates': trends['dates'],
        'cumulative_di': trends['cumulative'],
        'daily_new': trends['daily_new'],
        'daily_resolved': trends['daily_resolved'],
        'projects_data': projects_data,
        'project_names': project_names,
        'project_trends': trends['project_trends'],
        'milestones': milestones,
        'person_stats': person_stats,
        'version_plan': version_plan,
    })
    print('index.html generated successfully')
    print('path: ' + os.path.abspath(html_path))


if __name__ == '__main__':
    main()
