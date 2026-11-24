"""Automatic checks for "איפה הכסף". Uses only invented data.
Run from the repo folder:  python3 tests/run_tests.py
Needs: pip install playwright && playwright install chromium
"""
import functools, http.server, json, pathlib, sys, threading, socketserver
from playwright.sync_api import sync_playwright

ROOT = pathlib.Path(__file__).resolve().parent.parent
FAILS, PASSES = [], []

def check(name, cond, info=''):
    (PASSES if cond else FAILS).append(name)
    print(('  ✓ ' if cond else '  ✗ ') + name + ('' if cond else f'   → {info}'))

SETUP = """
D = blank(); D.cfg.start='2026-09-01'; D.cfg.people=['שי','שיראל']; D.cfg.me='שי';
D.cfg.pays=['מזומן','ביט / פייבוקס','עו"ש (הו"ק / העברות)','אשראי שי כאל'];
D.cfg.groups=[{id:'g1',name:'שוטף'},{id:'g2',name:'הכנסות'},{id:'g3',name:'דיור'}];
D.cfg.cats=[{id:'c1',name:'סופר',group:'g1',budget:0,kind:'variable'},{id:'c3',name:'מסעדות',group:'g1',budget:0,kind:'variable'},
  {id:'c10',name:'פארם',group:'g1',budget:0,kind:'variable'},
  {id:'c2',name:'משכורת',group:'g2',budget:14000,kind:'income'},{id:'c23',name:'הכנסות נוספות',group:'g2',budget:0,kind:'income'},
  {id:'c4',name:'משכנתא',group:'g3',budget:0,kind:'fixed'}];
D.kw=[['שופרסל','c1'],['סופר','c1'],['סופר פארם','c10'],['פיצה','c3'],['משכורת','c2']];
D.fixed=[{id:'f1',name:'משכנתא',cat:'c4',amount:4200,day:10,freq:'m',from:'2026-09'}];
save(false);
"""
CHAT = "\n".join([
    "05/11/2026, 10:15 - שיראל אשתי: 350 שופרסל אשראי שי כאל",
    "06/11/2026, 11:00 - שיראל אשתי: 80 מאפייה",
    "07/11/2026, 12:00 - שי: 120 סופר פארם",
    "08/11/2026, 09:00 - שי: 2 פיצות 90",
    "10/11/2026, 08:00 - שי: התקבל משכורת 14000",
])

def serve():
    class Quiet(http.server.SimpleHTTPRequestHandler):
        def log_message(self, *a): pass
    handler = functools.partial(Quiet, directory=str(ROOT))
    httpd = socketserver.TCPServer(('127.0.0.1', 0), handler)
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    return httpd

def main():
    httpd = serve(); url = f'http://127.0.0.1:{httpd.server_address[1]}/index.html'
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(viewport={'width': 390, 'height': 844}, has_touch=True)
        pg = ctx.new_page(); errs = []
        pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.goto(url); pg.wait_for_function('typeof render === "function" && document.querySelector("#app").innerHTML.length > 0')
        ev = pg.evaluate
        def imp(txt): return ev("async t => { await importFile(new File([t], 'c.txt')); stack.slice().forEach(e => closeSheet(e)); return D.chat.map(i => i.id); }", txt)

        print('versions')
        appv = ev('APP_VERSION'); swv = (ROOT / 'sw.js').read_text(encoding='utf-8'); vj = json.loads((ROOT / 'version.json').read_text(encoding='utf-8'))
        check('same version in app.js, sw.js and version.json', f"kesef-{appv}'" in swv and vj['version'] == appv, (appv, vj['version']))

        ev(SETUP)
        print('parser')
        ids = imp(CHAT)
        check('5 lines read', len(ids) == 5, ids)
        get = lambda part: ev("p => D.chat.find(i => i.text.includes(p))", part)
        check('longest keyword wins ("סופר פארם" → פארם, not סופר)', get('סופר פארם')['cat'] == 'c10', get('סופר פארם')['cat'])
        check('"התקבל משכורת" goes to the salary section', get('משכורת')['cat'] == 'c2', get('משכורת')['cat'])
        check('a line with two numbers is flagged', bool(get('פיצות').get('multi')), get('פיצות'))

        print('ids stay the same')
        bake = [i for i in ids if '80' in i.split('|')[1]][0]
        ev("id => { D.ov[id] = {cat: 'c3'}; save(false); }", bake)
        ids2 = imp(CHAT.replace('שיראל אשתי', 'שיראל ❤️'))
        check('contact renamed in the phone: same ids', sorted(ids) == sorted(ids2), (ids, ids2))
        check('contact renamed: chosen section kept', ev("id => (D.ov[id] || {}).cat", bake) == 'c3')
        ids3 = imp(CHAT.replace('80 מאפייה', '85 מאפייה <ההודעה נערכה>'))
        check('message edited in WhatsApp: same id', bake in ids3, ids3)
        check('edited message: new amount, section kept', ev("id => D.chat.find(i => i.id === id).amount", bake) == 85 and ev("id => (D.ov[id] || {}).cat", bake) == 'c3')
        check('no duplicate lines after re-imports', len(ids3) == 5, ids3)

        old = ev("() => { D.chat.forEach((i, n) => { const o = i.id; i.id = `${i.date}T${i.time}|${i.sender}|${i.amount}|old${n}#1`; if (D.ov[o]) { D.ov[i.id] = D.ov[o]; delete D.ov[o]; } }); save(false); return D.chat.map(i => i.id); }")
        ids4 = imp(CHAT.replace('80 מאפייה', '85 מאפייה <ההודעה נערכה>'))
        check('lines saved by older versions keep their ids (no duplicates in Excel)', sorted(old) == sorted(ids4), (old, ids4))

        print('export')
        res = ev("""async () => { let wbRef = null; window.download = () => {}; const X = await loadXLSX(); const w = X.write;
            X.write = (wb, o) => { wbRef = wb; return w(wb, o); }; await doExport('2026-11', '2026-11', true); X.write = w;
            const s = X.utils.sheet_to_json(wbRef.Sheets['סיכום'], {header: 1}); return s; }""")
        check('summary sheet has the sections', len(res) > 2 and any(r[1] == 'סופר' for r in res[1:]), res)

        print('backup')
        enc = ev("async () => { S.bkPw = 'secret123'; const t = await backupText(); return t; }")
        check('backup is encrypted by default', '"kesefBackup":2' in enc and 'שופרסל' not in enc)
        dec = ev("async t => { const o = JSON.parse(t); return JSON.parse(await decryptText(o, 'secret123')).chat.length; }", enc)
        check('encrypted backup opens with the password', dec == 5, dec)
        bad = ev("async t => { try { await decryptText(JSON.parse(t), 'wrong-pass'); return 'opened'; } catch { return 'refused'; } }", enc)
        check('wrong password is refused', bad == 'refused')

        print('restore')
        ev("async t => { const f = new File([t], 'b.json'); const o = JSON.parse(t); confirmRestore(JSON.parse(await decryptText(o, 'secret123')), f); }", enc)
        pg.wait_for_timeout(300)
        check('restore asks before replacing', pg.locator('#rs-go').count() == 1)
        ev("stack.slice().forEach(e => closeSheet(e))")

        print('storage')
        n = ev("async () => { await new Promise(r => setTimeout(r, 300)); return JSON.parse(await IDB.get('D')).chat.length; }")
        check('data is kept in the browser database (IndexedDB)', n == 5, n)

        print('undo')
        ev("lastTapAt = Date.now(); setOv(D.chat[0].id, {skip: true});")
        pg.wait_for_timeout(100)
        check('a change offers "ביטול"', pg.locator('#toast .undo').count() == 1)
        ev("doUndo()")
        check('undo brings it back', ev("!(D.ov[D.chat[0].id] || {}).skip"))

        print('excel file from the workbook')
        r = ev("""async () => { const X = await loadXLSX(); const mk = rows => X.utils.aoa_to_sheet(rows);
            const wb = X.utils.book_new();
            X.utils.book_append_sheet(wb, mk([['סעיף','קבוצה','סוג'],['סופר','שוטף','משתנה'],['מסעדות','שוטף','משתנה'],['פארם','שוטף','משתנה'],['משכורת','הכנסות','הכנסה'],['הכנסות נוספות','הכנסות','הכנסה'],['משכנתא','דיור','קבועה']]), 'סעיפים');
            X.utils.book_append_sheet(wb, mk([['מילה','סעיף'],['שופרסל','סופר']]), 'מילות מפתח');
            X.utils.book_append_sheet(wb, mk([['חודש','סעיף','סכום'],['9.26','סופר',3100],['9.26','משכנתא',4310]]), 'חודשים סגורים');
            X.utils.book_append_sheet(wb, mk([['שדה','ערך'],['פורמט',1]]), 'מידע');
            const buf = X.write(wb, {bookType: 'xlsx', type: 'array'});
            await importExcel(new File([buf], 'x.xlsx')); stack.slice().forEach(e => closeSheet(e));
            const ok1 = isClosed('2026-09') && monthData('2026-09').byCat['c4'] === 4310;
            const wb2 = X.utils.book_new(); X.utils.book_append_sheet(wb2, mk([['סעיף','קבוצה','סוג'],['סופר','שוטף','משתנה']]), 'סעיפים');
            X.utils.book_append_sheet(wb2, mk([['שדה','ערך'],['פורמט',99]]), 'מידע');
            const before = D.cfg.cats.length; await importExcel(new File([X.write(wb2, {bookType: 'xlsx', type: 'array'})], 'y.xlsx'));
            return {ok1, refusedNewer: D.cfg.cats.length === before}; }""")
        check('closed month amounts come from Excel', r['ok1'], r)
        check('a file from a newer Excel layout is refused, nothing changed', r['refusedNewer'], r)
        check('prototype of objects is frozen (crafted files can\'t pollute it)', ev("Object.isFrozen(Object.prototype)"))

        print('review')
        ev("S.month = '2026-11'; S.rv = 'todo'; ACT.view({dataset: {v: 'review'}})")
        check('review has the two segments', pg.locator('.rvseg button').count() == 2)
        check('multi-number line asks to check the amount', pg.locator('[data-act=mn-ok]').count() >= 1)
        ev("ACT.rvseg({dataset: {v: 'fixed'}})")
        check('fixed segment lists the fixed charges', 'משכנתא' in pg.locator('#app').inner_text())
        ev("ACT.rvseg({dataset: {v: 'todo'}})")

        print('home')
        ev("S.month = '2026-09'; ACT.view({dataset: {v: 'home'}})")
        txt = pg.locator('#app').inner_text()
        check('closed month says עודף/חוסר, not "צפויה"', ('עודף' in txt or 'חוסר' in txt) and 'יתרה צפויה' not in txt, txt[:300])

        print('data')
        ev("S.month = '2026-11'; S.f.sort = 'amt'; ACT.view({dataset: {v: 'data'}})")
        amts = ev("[...document.querySelectorAll('#app .txline .a')].map(e => parseFloat(e.textContent.replace(/[^\\d.]/g, '')))")
        check('sort by amount', amts == sorted(amts, reverse=True), amts)
        ev("S.f.sort = ''")

        print('version history')
        ev("historySheet()"); pg.wait_for_timeout(500)
        first = pg.locator('.sheet .hist').first.inner_text() if pg.locator('.sheet .hist').count() else ''
        check('history lists the installed version first', f'גרסה {appv}' in first and 'מותקנת' in first, first[:120])
        hist = json.loads((ROOT / 'changelog.json').read_text(encoding='utf-8'))
        check('changelog has the current version on top', hist[0]['version'] == appv, hist[0])
        ev("stack.slice().forEach(e => closeSheet(e))")

        print('screens render without errors')
        for v in ['home', 'data', 'review']:
            ev(f"ACT.view({{dataset: {{v: '{v}'}}}})")
        for m in ['2026-09', '2026-10', '2026-11']:
            ev(f"S.month = '{m}'; render()")
        check('no script errors', not errs, errs[:3])
        print('moving data from the old storage')
        ctx2 = b.new_context(); p2 = ctx2.new_page(); p2.goto(url); p2.wait_for_function('typeof render === "function"')
        p2.evaluate("async () => { const d = blank(); d.chat = [{id: 'x#1', date: '2026-11-01', time: '10:00', sender: 'a', person: 'a', text: '5 קפה', desc: 'קפה', amount: 5, cat: ''}]; await IDB.del('D'); localStorage.setItem(KEY, JSON.stringify(d)); }")
        p2.reload(); p2.wait_for_function('typeof render === "function" && D')
        moved = p2.evaluate("async () => ({n: D.chat.length, ls: localStorage.getItem(KEY), idb: !!(await IDB.get('D'))})")
        check('old data moved to the browser database, old copy removed', moved['n'] == 1 and moved['ls'] is None and moved['idb'], moved)
        ctx2.close()
        b.close()
    httpd.shutdown()
    print(f'\n{len(PASSES)} passed, {len(FAILS)} failed')
    sys.exit(1 if FAILS else 0)

if __name__ == '__main__':
    main()
