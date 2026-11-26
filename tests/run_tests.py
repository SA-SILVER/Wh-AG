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

def wait_js(page, expr, tries=100):
    # polls from outside the page: the page's security policy blocks Playwright's own in-page polling
    import time
    for _ in range(tries):
        try:
            if page.evaluate(expr): return
        except Exception: pass
        time.sleep(0.1)
    raise AssertionError('timeout: ' + expr)

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
        pg.goto(url); wait_js(pg, 'typeof render === "function" && document.querySelector("#app").innerHTML.length > 0')
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

        print('refunds')
        imp("\n".join([
            "03/11/2026, 10:00 - שי: 450 שח עבור השתלמות, צריך להתקבל החזר",
            "03/11/2026, 11:00 - שיראל: 1200 שח עבור משקפיים, יתקבל החזר מהביטוח",
            "04/11/2026, 11:00 - שי: 300 ציוד למשרד, צריך להתקבל החזר",
            "04/11/2026, 12:00 - שי: 200 מתנה לגננת, מגיע החזר מהוועד",
            "10/11/2026, 08:00 - שי: התקבל משכורת 14000",
            "12/11/2026, 09:00 - שי: התקבל 445 החזר",
            "15/11/2026, 09:00 - שיראל: התקבל 700 החזר משקפיים",
            "20/11/2026, 09:00 - שיראל: קיבלנו החזר 480 מהביטוח",
            "21/11/2026, 09:00 - שי: התקבל 150 החזר",
        ]))
        ev("() => { for (const it of D.chat) if (!it.incoming) D.ov[it.id] = {cat: /משקפיים/.test(it.text) ? 'c10' : /משכורת/.test(it.text) ? 'c2' : 'c3'}; D.ov[D.chat.find(i => /445/.test(i.text)).id] = {cat: 'c23'}; save(false); }")
        tx = lambda part: ev("p => { const t = chatTx(D.chat.find(i => i.text.includes(p))); return {amount: t.amount, cat: t.cat, done: t.refundDone, got: t.refundGot, ask: t.rfAsk, rfOf: t.rfOf}; }", part)
        r1 = tx('445')
        check('"צריך להתקבל החזר" marks the expense as waiting', tx('השתלמות')['got'] == 445, tx('השתלמות'))
        check('refund of a slightly different amount goes to the expense\'s section as a minus (even if it was put in income before)', r1['amount'] == -445 and r1['cat'] == 'c3', r1)
        check('…and closes that expense', tx('השתלמות')['done'], tx('השתלמות'))
        r2, r3 = tx('700'), tx('480')
        check('refund in two messages: both minus in the same section', r2['amount'] == -700 and r3['amount'] == -480 and r2['cat'] == r3['cat'] == 'c10', (r2, r3))
        check('expense closes only after most of it came back', tx('משקפיים')['got'] == 1180 and tx('משקפיים')['done'], tx('משקפיים'))
        check('"קיבלנו החזר" is money in', r3['amount'] < 0)
        r4 = tx('150')
        check('unclear refund is asked about, not guessed', r4['ask'] and not r4['rfOf'] and ev("rfAsks().length") == 1, r4)
        check('salary is not touched', tx('משכורת')['amount'] == 14000 and tx('משכורת')['cat'] == 'c2')
        ev("() => { const t = rfAsks()[0]; ACT['rf-link']({dataset: {id: t.id, e: D.chat.find(i => /גננת/.test(i.text)).id}}); }")
        r4 = tx('150')
        check('answering ties it: minus in that expense\'s section', r4['amount'] == -150 and r4['cat'] == 'c3' and ev("rfAsks().length") == 0, r4)
        ev("() => { const id = D.chat.find(i => i.text.includes('445')).id; ACT['rf-none']({dataset: {id}}); }"); pg.wait_for_timeout(300)
        allc = pg.locator('.sheet').last.locator('[data-pc]').evaluate_all("bs => bs.map(b => b.dataset.pc)")
        check('"לא קשור להוצאה" opens every section, income too', 'c23' in allc and 'c2' in allc and 'c1' in allc and '__skip' in allc, allc)
        check('…and nothing changes until a section is chosen', tx('445')['rfOf'], tx('445'))
        pg.locator('.sheet').last.locator('[data-pc=c1]').click(); pg.wait_for_timeout(300)
        check('choosing a section unties it and puts it there', not tx('445')['rfOf'] and tx('445')['cat'] == 'c1' and tx('445')['amount'] == -445, tx('445'))
        check('"not this expense" unties it', not tx('445')['rfOf'] and not tx('השתלמות')['done'], (tx('445'), tx('השתלמות')))
        ev("() => { const id = D.chat.find(i => i.text.includes('445')).id; ACT['t-reset']({dataset: {id}}); }")
        check('an expense and its refund on the same day are not called a double entry', ev("dupPairs().every(([a, b]) => !!a.incoming === !!b.incoming)"))
        ev("S.month = '2026-11'; S.rv = 'todo'; ACT.view({dataset: {v: 'review'}})")
        check('review shows what is still waiting for a refund', 'מחכות להחזר' in pg.locator('#app').inner_text())

        print('refund for an expense from before the budget start')
        imp("\n".join([
            "15/08/2026, 10:00 - שי: 380 שח עבור קורס עזרה ראשונה, צריך להתקבל החזר",
            "12/10/2026, 09:00 - שי: התקבל 380 החזר",
        ]))
        pre = ev("() => { const t = chatTx(D.chat.find(i => /380 החזר/.test(i.text))); return {amount: t.amount, cat: t.cat, rfPre: t.rfPre, pending: t.pending, inPend: pendingChat().some(x => x.id === t.id), inChat: D.chat.some(i => /קורס/.test(i.text))}; }")
        check('expense before the start is not counted, but kept aside', not pre['inChat'] and ev("(D.rfPre || []).length") == 1, pre)
        check('the refund is tied to it and asks for the expense section (not income)', pre['rfPre'] and pre['pending'] and pre['inPend'], pre)
        ev("() => { const g = pendGroups().find(g => g.items.some(t => t.rfPre)); groupPick(g.key, 'c3'); }")
        pre = ev("() => { const t = chatTx(D.chat.find(i => /380 החזר/.test(i.text))); const e = D.rfPre[0]; return {amount: t.amount, cat: t.cat, expCat: (D.ov[e.id] || {}).cat}; }")
        check('…then it comes off that expense section as a minus', pre['amount'] == -380 and pre['cat'] == 'c3' and pre['expCat'] == 'c3', pre)
        check('…and is not in the Excel export of the month before the start', ev("monthTx('2026-08').length") == 0)

        print('money in without "החזר"')
        imp("\n".join([
            "01/12/2026, 10:00 - שי: 260 שח ספרים לקורס, צריך להתקבל החזר",
            "05/12/2026, 10:00 - שי: התקבל 260",
            "06/12/2026, 10:00 - שי: התקבל 500 מתנה",
        ]))
        g = lambda part: ev("p => { const t = chatTx(D.chat.find(i => i.text.includes(p))); return {amount: t.amount, cat: t.cat, rfOf: !!t.rfOf, pending: t.pending, inc: isIncome(catById(t.cat))}; }", part)
        check('"התקבל 260" (no "החזר") comes off the expense that waited', g('התקבל 260')['rfOf'] and g('התקבל 260')['amount'] == -260, g('התקבל 260'))
        check('"התקבל" with no income word is not counted as income – asks for a section', g('500 מתנה')['pending'] and not g('500 מתנה')['inc'], g('500 מתנה'))
        qc = ev("() => quickCats(chatTx(D.chat.find(i => /500 מתנה/.test(i.text)))).map(c => [c.id, isIncome(c)])")
        check('…offering expense sections first and the income sections too, in one tap', qc and not qc[0][1] and any(x[1] for x in qc) and any(x[0] == 'c23' for x in qc), qc)
        ev("() => { const g = pendGroups().find(g => g.items.some(t => /500 מתנה/.test(t.text))); groupPick(g.key, 'c23'); }")
        check('picking an income section counts it as income', g('500 מתנה')['amount'] == 500 and g('500 מתנה')['cat'] == 'c23' and not g('500 מתנה')['pending'], g('500 מתנה'))
        ev("() => { const id = D.chat.find(i => /500 מתנה/.test(i.text)).id; delete D.ov[id]; D.kw = D.kw.filter(([k]) => k !== 'מתנה'); recategorize(); save(false); }")
        check('"התקבל משכורת" stays income', g('משכורת')['cat'] == 'c2' and g('משכורת')['amount'] == 14000)
        retro = ev("() => { const it = D.chat.find(i => /500 מתנה/.test(i.text)); it.cat = 'c23'; D.v = 5; migrate6(); return {cat: it.cat, v: D.v}; }")
        check('applied to lines already in the app (old "הכנסות נוספות" removed)', retro['cat'] == '' and retro['v'] == 6, retro)
        ev("() => { const id = D.chat.find(i => /500 מתנה/.test(i.text)).id; D.ov[id] = {cat: 'c23'}; save(false); }")
        check('income chosen by hand is kept', g('500 מתנה')['cat'] == 'c23' and g('500 מתנה')['amount'] == 500)

        print('income detail of a closed month')
        ev("() => { D.closed = D.closed || {}; D.closed['2026-11'] = {c2: 14000, c23: 2500}; save(false); S.month = '2026-11'; }")
        ev("incDetailSheet()"); pg.wait_for_timeout(300)
        it = pg.locator('.sheet').last.inner_text()
        check('closed month: each income section shows the lines it is made of', 'הכנסות נוספות' in it and 'התקבל משכורת' in it and 'אין פירוט באפליקציה' in it, it[:400])
        ev("stack.slice().forEach(e => closeSheet(e)); delete D.closed['2026-11']; save(false);"); pg.wait_for_timeout(250)

        print('money in, in a month closed in Excel')
        ev("() => { const id = D.chat.find(i => /500 מתנה/.test(i.text)).id; delete D.ov[id]; D.kw = D.kw.filter(([k]) => k !== 'מתנה'); recategorize(); D.closed = D.closed || {}; D.closed['2026-12'] = {c1: 100}; save(false); }")
        check('a "התקבל" line without a section in a closed month is asked about in review', ev("pendingChat().some(t => /500 מתנה/.test(t.text))"))
        ev("S.month = '2026-12'; S.rv = 'todo'; ACT.view({dataset: {v: 'review'}})")
        # the explanation is folded behind ⓘ – it is in the page, opened on a tap
        check('…marked as a closed month, with what to do next', 'חודש סגור' in pg.locator('#app').inner_text() and 'ייבוא מאיפה הכסף' in pg.locator('#app').text_content())
        ev("() => { const g = pendGroups().find(g => g.items.some(t => /500 מתנה/.test(t.text))); groupPick(g.key, 'c1'); }")
        ex = ev("""async () => { let wbRef = null; window.download = () => {}; const X = await loadXLSX(); const w = X.write;
            X.write = (wb, o) => { wbRef = wb; return w(wb, o); }; await doExport('2026-12', '2026-12', true); X.write = w;
            return X.utils.sheet_to_json(wbRef.Sheets['תנועות'], {header: 1}).filter(r => String(r[7] || '').includes('500 מתנה')).map(r => [r[3], r[5]]); }""")
        check('…and once it has a section it goes out in the Excel export of that month', ex == [['סופר', -500]], ex)
        ev("() => { delete D.closed['2026-12']; const id = D.chat.find(i => /500 מתנה/.test(i.text)).id; delete D.ov[id]; D.kw = D.kw.filter(([k]) => k !== 'מתנה'); recategorize(); save(false); }")

        print('"התקבל" with no amount (a reply in the group)')
        na0 = ev("() => { const L = refundLinks(); return Object.keys(L.byId).filter(id => !rfState(id).done); }")
        imp("\n".join([
            "20/12/2026, 10:00 - שי: התקבל",
            "03/01/2027, 10:00 - שי: 400 שח עבור לוחות שנה, צריך להתקבל החזר של 375",
            "10/01/2027, 10:00 - שיראל: התקבל ונמשך לבנק",
            "11/01/2027, 10:00 - שי: התקבל",
        ]))
        ev("() => { const e = D.chat.find(i => /לוחות שנה/.test(i.text)); D.ov[e.id] = {cat: 'c3'}; save(false); }")
        first = ev("() => { const t = chatTx(D.chat.find(i => i.text === 'התקבל ונמשך לבנק')); return t.rfAsk ? t.rfAsk[0] : 'linked'; }")
        check('with several expenses waiting it asks, newest first', first == ev("() => D.chat.find(i => /לוחות שנה/.test(i.text)).id"), first)
        ev("ids => { for (const id of ids) D.ov[id] = {...(D.ov[id] || {}), skip: true}; save(false); }", na0)
        na = ev("() => D.chat.filter(i => /^התקבל/.test(i.text)).map(i => { const t = chatTx(i); return {text: i.text, amount: t.amount, cat: t.cat, rfOf: !!t.rfOf, ask: t.rfAsk}; })")
        na = [x for x in na if x['text'] in ('התקבל ונמשך לבנק', 'התקבל')]
        check('"התקבל" lines without an amount are read, not dropped', len(na) == 3, na)
        nb = [x for x in na if x['text'] == 'התקבל ונמשך לבנק']
        check('…tied to the expense that waits, for the amount it waits for (375), as a minus in its section', nb and nb[0]['rfOf'] and nb[0]['amount'] == -375 and nb[0]['cat'] == 'c3', nb)
        check('…and the expense is closed', ev("() => chatTx(D.chat.find(i => /לוחות שנה/.test(i.text))).refundDone"))
        check('no amount and nothing waiting: asked in review, not guessed', any(x['text'] == 'התקבל' and x['ask'] == [] for x in na) and ev("rfAsks().some(t => t.noAmt)"), na)
        check('two "התקבל" without amount are not called a double entry', ev("dupPairs().every(([a, b]) => !/^התקבל$/.test(a.text) || !/^התקבל/.test(b.text))"))
        ev("() => D.chat.filter(i => i.text === 'התקבל').forEach(i => ACT['rf-skip']({dataset: {id: i.id}}))")
        check('"להתעלם" removes it from review', not ev("rfAsks().some(t => t.noAmt)"))
        ev("ids => { for (const id of ids) { delete D.ov[id].skip; } save(false); }", na0)

        print('home squares')
        ev("S.month = '2026-11'; ACT.view({dataset: {v: 'home'}})")
        pg.locator('[data-act=incdet]').click(); pg.wait_for_timeout(300)
        stxt = pg.locator('.sheet').last.inner_text() if pg.locator('.sheet').count() else ''
        check('tapping "הכנסות" opens the income detail', 'הכנסות –' in stxt and 'משכורת' in stxt, stxt[:200])
        ev("stack.slice().forEach(e => closeSheet(e))"); pg.wait_for_timeout(250)
        ev("S.month = '2026-10'; ACT.view({dataset: {v: 'home'}})")
        pg.locator('.kpi [data-rv=fixed]').click(); pg.wait_for_timeout(250)
        check('tapping "מתוכן קבועות" goes to the fixed charges', ev("S.view") == 'review' and ev("S.rv") == 'fixed')

        ev("S.month = '2026-11'; ACT.view({dataset: {v: 'home'}})")
        pg.locator('[data-act=baldet]').click(); pg.wait_for_timeout(300)
        btxt = pg.locator('.sheet').last.inner_text() if pg.locator('.sheet').count() else ''
        check('tapping the balance shows how it is made (income − out − fixed still to come)', 'הכנסות' in btxt and 'קבועות שעוד צפויות' in btxt and 'לא נכללות' in btxt, btxt[:300])
        ev("stack.slice().forEach(e => closeSheet(e))"); pg.wait_for_timeout(250)
        sel = ev("({body: getComputedStyle(document.body).userSelect, inp: (() => { const i = document.createElement('input'); document.body.appendChild(i); const v = getComputedStyle(i).userSelect; i.remove(); return v; })()})")
        check('text on the screen can\'t be selected, typing fields can', sel['body'] == 'none' and sel['inp'] != 'none', sel)

        print('swipe between months')
        SW = """([sel, dx]) => { const el = sel === 'sheet' ? stack[stack.length - 1].sh.querySelector('.chip, h3, .card') : document.querySelector(sel);
            const r = el.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + Math.min(r.height / 2, 20);
            const T = (cx) => new Touch({identifier: 1, target: el, clientX: cx, clientY: y});
            el.dispatchEvent(new TouchEvent('touchstart', {bubbles: true, touches: [T(x)], changedTouches: [T(x)]}));
            el.dispatchEvent(new TouchEvent('touchend', {bubbles: true, touches: [], changedTouches: [T(x + dx)]}));
            return S.month; }"""
        ev("S.month = '2026-10'; S.rv = 'todo'; ACT.view({dataset: {v: 'review'}})")
        check('swipe right = next month (starting on a choice button too)', ev(SW, ['.chips button', 120]) == '2026-11')
        check('swipe left = previous month', ev(SW, ['#app .hero', -120]) == '2026-10')
        ev("ACT.view({dataset: {v: 'data'}})")
        check('swipe works on the details screen', ev(SW, ['#app .hero', 120]) == '2026-11')
        ev("groupSheet(D.cfg.groups[0].id)"); pg.wait_for_timeout(300)
        m = ev(SW, ['sheet', -120]); pg.wait_for_timeout(250)
        check('swipe works inside a window that shows a month, and the window follows', m == '2026-10' and 'אוקטובר' in pg.locator('.sheet').last.inner_text(), (m, pg.locator('.sheet').last.inner_text()[:80]))
        ev("stack.slice().forEach(e => closeSheet(e))"); pg.wait_for_timeout(250)
        ev("txSheet(D.chat[0].id)"); pg.wait_for_timeout(300)
        check('a window of one line does not change month', ev(SW, ['sheet', 120]) == '2026-10')
        ev("stack.slice().forEach(e => closeSheet(e))"); pg.wait_for_timeout(250)

        print('fits the phone')
        vp = pg.locator('meta[name=viewport]').get_attribute('content')
        check('zoom is turned off', 'user-scalable=no' in vp and 'maximum-scale=1' in vp, vp)
        pg.set_viewport_size({'width': 360, 'height': 760})
        wide = ev("""async () => { const out = []; const chk = n => { const w = document.documentElement.scrollWidth; if (w > innerWidth) out.push(n + ':' + w); };
            for (const m of ['2026-09', '2026-10', '2026-11']) for (const v of ['home', 'data', 'review']) { S.month = m; S.rv = 'todo'; ACT.view({dataset: {v}}); chk(v + m); }
            S.rv = 'fixed'; render(); chk('fixed');
            S.month = '2026-11'; ACT.view({dataset: {v: 'home'}}); incDetailSheet(); await new Promise(r => setTimeout(r, 300)); chk('incdet');
            stack.slice().forEach(e => closeSheet(e)); txSheet(D.chat.find(i => i.incoming).id); await new Promise(r => setTimeout(r, 300)); chk('tx');
            stack.slice().forEach(e => closeSheet(e)); return out; }""")
        check('no screen is wider than a 360px phone', not wide, wide)
        pg.set_viewport_size({'width': 390, 'height': 844})

        print('fixed expenses: clear parts')
        r = ev("""() => { S.month = curMonth(); const html = fixedByGroup(S.month); const d = dash(S.month);
            const rows = fixedRows(S.month); const exp = rows.filter(t => t.st === 'exp').reduce((s, t) => s + t.amount, 0);
            return {sameAsHome: Math.abs(exp - d.upc) < 0.01, hasParts: html.includes('כבר ירד') && html.includes('עוד צפוי לרדת')}; }""")
        check('"עוד צפוי לרדת" in the fixed list = the home screen number', r['sameAsHome'], r)
        check('the fixed list shows "כבר ירד" and "עוד צפוי לרדת"', r['hasParts'], r)
        r2 = ev("""() => { S.rv = 'todo'; S.view = 'review'; render(); const t = document.querySelector('#app').innerText;
            return {noConfirm: !document.querySelector('#app [data-act=fx-ok]'), note: t.includes('עוד לא נסגר') || t.includes('עוד לא נסגרו')}; }""")
        check('"לטיפול" has no "ירד ✓" buttons for fixed charges and incomes', r2['noConfirm'], r2)
        check('"לטיפול" notes the months not closed in Excel', r2['note'], r2)

        print('version 3.4: review, home, data, menu')
        ev("stack.slice().forEach(e => closeSheet(e)); S.month = '2026-11'; S.rv = 'todo'; ACT.view({dataset: {v: 'review'}})")
        seg = ev("(document.querySelector('.rvseg button').textContent.match(/\\((\\d+)\\)/) || [0, '0'])[1]")
        dot = ev("(document.querySelector('nav .dot') || {}).textContent || '0'")
        check('the badge and the "לטיפול" number are the same', seg == dot, (seg, dot))
        check('the badge counts choices, not lines', ev("reviewCount() < pendingChat().length + rfAsks().length + dupPairs().length + 1 || pendGroups().length === pendingChat().length"))
        ev("""async () => { await importFile(new File(['24/11/2026, 09:00 - שי: 3 חולצות 120'], 'm.txt')); stack.slice().forEach(e => closeSheet(e)); S.rv = 'todo'; S.view = 'review'; render(); }""")
        mid = ev("D.chat.find(i => /3 חולצות 120/.test(i.text)).id")
        cards = ev("id => [...document.querySelectorAll('#app .card.rv')].filter(c => c.textContent.includes('3 חולצות 120')).length", mid)
        check('a line with two numbers and no section shows up once', cards == 1, cards)
        check('…asking for the amount first, one chip per number', ev("[...document.querySelectorAll('[data-act=mn-set]')].some(b => b.dataset.v === '120')"))
        ev("id => ACT['mn-set']({dataset: {id, v: '120'}})", mid)
        t = ev("id => { const t = chatTx(D.chat.find(i => i.id === id)); return {a: t.amount, open: multiOpen(t)}; }", mid)
        check('…choosing 120 sets the amount, then the section is asked', t['a'] == 120 and not t['open'] and ev("id => !!document.querySelector('[data-act=rvg-cat]')", mid), t)
        qc = ev("() => quickCats({desc: 'קפה', text: 'קפה', incoming: false}).map(c => c.id)")
        check('a fixed-type section comes only after the ordinary ones in the quick choices', ('c4' not in qc or qc.index('c4') == len(qc) - 1) and len(qc) <= 4, qc)
        ev("stack.slice().forEach(e => closeSheet(e)); ACT.add()"); pg.wait_for_timeout(250)
        check('"+" opens the manual entry straight away', pg.locator('#m-amt').count() == 1)
        ev("stack.slice().forEach(e => closeSheet(e)); settingsSheet()"); pg.wait_for_timeout(250)
        check('the menu is in groups', pg.locator('.sheet .gname').count() == 4)
        ev("stack.slice().forEach(e => closeSheet(e)); S.month = curMonth(); ACT.view({dataset: {v: 'home'}})")
        check('home: fixed charges with no payment method are not a payment row', 'קבועות (בלי אמצעי תשלום)' not in pg.locator('#app').inner_text())
        ev("D.meta.lastImport = {...(D.meta.lastImport || {}), at: Date.now() - 5 * 864e5}; S.dismissed.delete('stale'); render()")
        check('home: an old import is pointed out', pg.locator('.banner.stale').count() == 1)
        ev("D.meta.lastImport.at = Date.now(); render()")
        check('…and not when it is fresh', pg.locator('.banner.stale').count() == 0)
        check('the section is not repeated when it is the description', ev("descCat({desc: 'משכנתא'}, {name: 'משכנתא'}) === 'משכנתא' && descCat({desc: 'שופרסל'}, {name: 'סופר'}) === 'שופרסל · סופר'"))
        ev("ACT.view({dataset: {v: 'data'}})")
        check('data: the filters sit behind one button', pg.locator('[data-fsel]').count() == 0 and pg.locator('[data-act=fopen]').count() == 1)
        ev("ACT.fopen()")
        check('…and open on a tap', pg.locator('[data-fsel]').count() == 3)
        ev("ACT.fopen()")
        print('screens render without errors')
        for v in ['home', 'data', 'review']:
            ev(f"ACT.view({{dataset: {{v: '{v}'}}}})")
        for m in ['2026-09', '2026-10', '2026-11']:
            ev(f"S.month = '{m}'; render()")
        check('no script errors', not errs, errs[:3])
        print('wipe keeps the section settings')
        nc = ev("D.cfg.cats.length"); nk = ev("D.kw.length")
        ev("stack.slice().forEach(e => closeSheet(e)); wipeSheet()"); pg.wait_for_timeout(250)
        pg.click('#w-go'); pg.wait_for_timeout(300)
        pg.fill('#pw1', 'keep-pass-1'); pg.fill('#pw2', 'keep-pass-1'); pg.click('#pw-ok'); pg.wait_for_timeout(1500)
        st = ev("async () => { const k = await IDB.get('keep'); return {D: D === null, keep: !!k, plain: k ? (k.includes('סופר') || k.includes('שופרסל')) : true, btn: !!document.querySelector('[data-act=keep-restore]')}; }")
        check('wipe deleted the data', st['D'], st)
        check('section settings kept, encrypted (no readable names)', st['keep'] and not st['plain'], st)
        check('the opening screen offers to bring the settings back', st['btn'], st)
        pg.click('[data-act=keep-restore]'); pg.wait_for_timeout(250)
        pg.fill('#pw1', 'keep-pass-1'); pg.click('#pw-ok'); pg.wait_for_timeout(1500)
        r = ev("async () => ({cats: D ? D.cfg.cats.length : -1, kw: D ? D.kw.length : -1, chat: D ? D.chat.length : -1, money: D ? D.cfg.cats.some(c => num(c.budget) > 0) : true, keepLeft: !!(await IDB.get('keep'))})")
        check('settings came back: same sections and keywords', r['cats'] == nc and r['kw'] == nk, (r, nc, nk))
        check('no amounts came back (no lines, no expected income)', r['chat'] == 0 and not r['money'], r)
        check('the kept copy is removed after use', not r['keepLeft'], r)
        ev(SETUP)

        print('moving data from the old storage')
        ctx2 = b.new_context(); p2 = ctx2.new_page(); p2.goto(url); wait_js(p2, 'typeof render === "function"')
        p2.evaluate("async () => { const d = blank(); d.chat = [{id: 'x#1', date: '2026-11-01', time: '10:00', sender: 'a', person: 'a', text: '5 קפה', desc: 'קפה', amount: 5, cat: ''}]; await IDB.del('D'); localStorage.setItem(KEY, JSON.stringify(d)); }")
        p2.reload(); wait_js(p2, 'typeof render === "function" && !!D')
        moved = p2.evaluate("async () => ({n: D.chat.length, ls: localStorage.getItem(KEY), idb: !!(await IDB.get('D'))})")
        check('old data moved to the browser database, old copy removed', moved['n'] == 1 and moved['ls'] is None and moved['idb'], moved)
        ctx2.close()
        b.close()
    httpd.shutdown()
    print(f'\n{len(PASSES)} passed, {len(FAILS)} failed')
    sys.exit(1 if FAILS else 0)

if __name__ == '__main__':
    main()
