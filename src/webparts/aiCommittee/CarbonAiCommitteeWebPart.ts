import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

const SP_LIST = 'AI_Committee';

const PC: Record<string, {num:string;cls:string;bar:string}> = {
  p1:{num:'Pillar 1',cls:'p1',bar:'#185fa5'},
  p2:{num:'Pillar 2',cls:'p2',bar:'#0f6e56'},
  p3:{num:'Pillar 3',cls:'p3',bar:'#7b3fa0'},
  p4:{num:'Pillar 4',cls:'p4',bar:'#ba7517'}
};

interface ITask    { id:string; text:string; done:boolean; spId?:number; }
interface IProject { id:string; title:string; phase:string; tasks:ITask[]; }
interface IPillar  { id:string; title:string; champion:string; desc:string; projects:IProject[]; }
interface IAgenda  { num:string; time:string; title:string; tag:string; tagLabel:string; points:string[]; note:string; }

export default class CarbonAiCommitteeWebPart extends BaseClientSideWebPart<Record<string,never>> {

  private _editMode    = false;
  private _activeTab   = 'overview';
  private _spConnected = false;
  private _stTimer: any= null;
  private _bound       = false;

  private _p: IPillar[] = [];
  private _a: IAgenda[] = [];

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  // ── LIFECYCLE ──────────────────────────────────────────────────────────────
  public render(): void {
    this._p = this._defaultPillars();
    this._a = this._defaultAgenda();
    this.domElement.innerHTML = this._css() + this._shell();
    this._renderAll();
    if (!this._bound) { this._bind(); this._bound = true; }
    this._connect();
  }

  // ── SHELL & CSS ───────────────────────────────────────────────────────────
  private _shell(): string {
    return `
<div class="cg-hdr">
  <div class="cg-hdr-i">
    <div>
      <div class="cg-h1">Carbon Group &mdash; AI Committee</div>
      <div class="cg-sub">4 Pillars &middot; Monthly Meetings &middot; Living Roadmap</div>
    </div>
    <div id="cgSt" class="cg-st cg-off">&#9675; Local</div>
    <div class="cg-hbtns">
      <span class="cg-epill" id="cgEpill">Editing</span>
      <button class="cg-btn cg-bw" id="cgSync">&#8635; Sync</button>
      <button class="cg-btn cg-bw" id="cgEdit"><span id="cgELbl">&#9998; Enable editing</span></button>
      <button class="cg-btn cg-bg" id="cgExp">&#11015; Export HTML</button>
    </div>
  </div>
</div>
<div class="cg-main">
  <div class="cg-tabs">
    <button class="cg-tab cg-ta" id="cgt-overview" data-tab="overview">Overview &amp; Pillars</button>
    <button class="cg-tab" id="cgt-agenda"   data-tab="agenda">Monthly Agenda</button>
    <button class="cg-tab" id="cgt-tasks"    data-tab="tasks">Tasks by Pillar</button>
  </div>
  <div id="cgv-overview"></div>
  <div id="cgv-agenda"   class="cg-hide"></div>
  <div id="cgv-tasks"    class="cg-hide"></div>
  <div class="cg-foot"><strong>Document owner:</strong> Nathan (Carbon Group) &nbsp;&middot;&nbsp;
    <strong>Cadence:</strong> Monthly &nbsp;&middot;&nbsp;
    <strong>Questions:</strong> Raise with Nathan before proceeding</div>
</div>`;
  }

  private _css(): string {
    return `<style>
.cg-wrap *{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#fff;--bg2:#f8f7f4;--bg3:#f1efe8;--tx:#1a1a18;--tx2:#5f5e5a;--tx3:#888780;
  --bd:#e0ddd5;--bd2:#ccc9c0;--navy:#042c53;--red:#a32d2d;--redl:#fcebeb;--redb:#f09595;
  --p1:#185fa5;--p1l:#e6f1fb;--p1b:#85b7eb;
  --p2:#0f6e56;--p2l:#e1f5ee;--p2b:#5dcaa5;
  --p3:#7b3fa0;--p3l:#f3eafe;--p3b:#c49de0;
  --p4:#ba7517;--p4l:#faeeda;--p4b:#fac775;
  --r:10px;--sh:0 1px 4px rgba(0,0,0,.07)}
.cg-hdr{background:var(--navy);padding:18px 0;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.18);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}
.cg-hdr-i{max-width:900px;margin:0 auto;padding:0 24px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}
.cg-h1{font-size:20px;font-weight:500;color:#fff;line-height:1.2}
.cg-sub{font-size:13px;color:#93b8d8;margin-top:3px}
.cg-hbtns{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.cg-btn{display:inline-flex;align-items:center;gap:5px;padding:8px 14px;border-radius:8px;font-size:13px;font-weight:500;cursor:pointer;border:1px solid transparent;transition:all .15s;font-family:inherit;background:none}
.cg-bw{background:rgba(255,255,255,.12);color:#fff;border-color:rgba(255,255,255,.25)}
.cg-bw:hover{background:rgba(255,255,255,.22)}
.cg-bg{background:var(--p2);color:#fff;border-color:var(--p2)}
.cg-bg:hover{filter:brightness(.9)}
.cg-bsm{padding:5px 10px;font-size:12px}
.cg-bdanger{background:var(--redl);color:var(--red);border-color:var(--redb)}
.cg-epill{background:var(--p4l);color:var(--p4);border:1px solid var(--p4b);border-radius:20px;padding:4px 12px;font-size:12px;font-weight:500;display:none}
.cg-epill.on{display:inline-block}
.cg-st{font-size:11px;padding:4px 12px;border-radius:20px;font-weight:500;white-space:nowrap;border:1px solid transparent;letter-spacing:.01em;transition:all .4s}
.cg-con{background:rgba(93,202,165,.12);color:#5dcaa5;border-color:rgba(93,202,165,.3)}
.cg-sav{background:rgba(93,202,165,.2);color:#5dcaa5;border-color:rgba(93,202,165,.4)}
.cg-sng{background:rgba(24,95,165,.15);color:#85b7eb;border-color:rgba(133,183,235,.3)}
.cg-cng{background:rgba(255,255,255,.08);color:#fac775;border-color:rgba(250,199,117,.2)}
.cg-off{background:rgba(255,255,255,.06);color:#93b8d8;border-color:rgba(255,255,255,.1)}
.cg-err{background:rgba(163,45,45,.15);color:#f09595;border-color:rgba(163,45,45,.25)}
.cg-main{max-width:900px;margin:0 auto;padding:28px 24px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:var(--tx)}
.cg-tabs{display:flex;gap:4px;margin-bottom:24px;background:var(--bg3);border:1px solid var(--bd);border-radius:var(--r);padding:4px}
.cg-tab{flex:1;padding:10px 8px;border:none;background:none;border-radius:8px;font-size:13px;font-weight:500;color:var(--tx2);cursor:pointer;transition:all .15s;font-family:inherit;text-align:center}
.cg-ta{background:#fff;color:var(--tx);box-shadow:var(--sh)}
.cg-hide{display:none}
.cg-ibox{background:#fff;border:1px solid var(--bd);border-radius:var(--r);padding:18px 20px;margin-bottom:24px;line-height:1.7}
.cg-ibox p{font-size:13px;color:var(--tx2);margin-bottom:8px}
.cg-ibox p:last-child{margin-bottom:0}
.cg-ibox strong{color:var(--tx)}
.cg-pgw{background:#fff;border:1px solid var(--bd);border-radius:var(--r);padding:16px 20px;margin-bottom:24px}
.cg-pgt{font-size:14px;font-weight:500;margin-bottom:14px}
.cg-pgr{display:flex;align-items:center;gap:12px;margin-bottom:10px}
.cg-pgr:last-child{margin-bottom:0}
.cg-pgl{font-size:12px;font-weight:500;width:210px;flex-shrink:0}
.cg-pgb{flex:1;height:8px;border-radius:4px;background:var(--bg3);overflow:hidden}
.cg-pgf{height:100%;border-radius:4px;transition:width .4s}
.cg-pgp{font-size:12px;color:var(--tx3);width:80px;text-align:right;flex-shrink:0}
.cg-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px}
@media(max-width:620px){.cg-grid{grid-template-columns:1fr}}
.cg-pc{background:#fff;border-radius:var(--r);border:2px solid var(--bd);overflow:hidden;box-shadow:var(--sh)}
.cg-ph{padding:16px 18px;border-bottom:1px solid var(--bd)}
.cg-pn{font-size:11px;font-weight:600;letter-spacing:.07em;text-transform:uppercase;margin-bottom:6px}
.cg-pt{font-size:16px;font-weight:500;margin-bottom:6px}
.cg-pch{font-size:12px;color:var(--tx2)}
.cg-pch span{font-weight:500}
.cg-chin{font-size:12px;padding:4px 8px;border:1px solid var(--bd2);border-radius:6px;font-family:inherit;background:var(--bg2);color:var(--tx);width:180px}
.cg-pb{padding:14px 18px}
.cg-pd{font-size:13px;color:var(--tx2);line-height:1.6;margin-bottom:12px}
.cg-pd textarea{width:100%;min-height:70px;font-size:13px;border:1px solid var(--bd2);border-radius:6px;padding:7px;resize:vertical;font-family:inherit;color:var(--tx);background:var(--bg2)}
.cg-tc{font-size:12px;font-weight:500;padding:3px 10px;border-radius:12px;display:inline-block}
.cg-p1 .cg-ph{background:var(--p1l);border-bottom-color:var(--p1b)}.cg-p1 .cg-pn,.cg-p1 .cg-pt{color:var(--p1)}.cg-p1 .cg-tc{background:var(--p1l);color:var(--p1)}.cg-p1{border-color:var(--p1b)}
.cg-p2 .cg-ph{background:var(--p2l);border-bottom-color:var(--p2b)}.cg-p2 .cg-pn,.cg-p2 .cg-pt{color:var(--p2)}.cg-p2 .cg-tc{background:var(--p2l);color:var(--p2)}.cg-p2{border-color:var(--p2b)}
.cg-p3 .cg-ph{background:var(--p3l);border-bottom-color:var(--p3b)}.cg-p3 .cg-pn,.cg-p3 .cg-pt{color:var(--p3)}.cg-p3 .cg-tc{background:var(--p3l);color:var(--p3)}.cg-p3{border-color:var(--p3b)}
.cg-p4 .cg-ph{background:var(--p4l);border-bottom-color:var(--p4b)}.cg-p4 .cg-pn,.cg-p4 .cg-pt{color:var(--p4)}.cg-p4 .cg-tc{background:var(--p4l);color:var(--p4)}.cg-p4{border-color:var(--p4b)}
.cg-aw{background:#fff;border:1px solid var(--bd);border-radius:var(--r);overflow:hidden;margin-bottom:20px;box-shadow:var(--sh)}
.cg-ah{padding:16px 20px;border-bottom:1px solid var(--bd);display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.cg-ah h2{font-size:16px;font-weight:500;flex:1}
.cg-am{font-size:12px;color:var(--tx3)}
.cg-ai{border-bottom:1px solid var(--bd)}
.cg-ai:last-child{border-bottom:none}
.cg-aih{padding:14px 20px;display:flex;align-items:center;gap:12px;cursor:pointer;transition:background .12s}
.cg-aih:hover{background:var(--bg2)}
.cg-anum{font-size:13px;font-weight:600;width:28px;flex-shrink:0}
.cg-atm{font-size:12px;color:var(--tx3);width:56px;flex-shrink:0}
.cg-atl{font-size:14px;font-weight:500;flex:1}
.cg-atg{font-size:11px;font-weight:500;padding:2px 9px;border-radius:12px;border:1px solid transparent;white-space:nowrap}
.cg-atog{font-size:11px;color:var(--tx3);white-space:nowrap}
.cg-ab{padding:14px 20px 18px 60px;display:none;border-top:1px solid var(--bd);background:var(--bg2)}
.cg-ab.open{display:block}
.cg-ab ul{padding-left:18px}
.cg-ab li{font-size:13px;color:var(--tx2);margin-bottom:6px;line-height:1.6}
.cg-ab .cg-note{font-size:12px;color:var(--tx3);margin-top:10px;font-style:italic}
.tg-all{background:var(--bg3);color:var(--tx2);border-color:var(--bd2)}
.tg-p1{background:var(--p1l);color:var(--p1);border-color:var(--p1b)}
.tg-p2{background:var(--p2l);color:var(--p2);border-color:var(--p2b)}
.tg-p3{background:var(--p3l);color:var(--p3);border-color:var(--p3b)}
.tg-p4{background:var(--p4l);color:var(--p4);border-color:var(--p4b)}
.cg-ps{background:#fff;border-radius:var(--r);border:1px solid var(--bd);overflow:hidden;margin-bottom:16px;box-shadow:var(--sh)}
.cg-psh{padding:14px 18px;border-bottom:1px solid var(--bd);display:flex;align-items:center;gap:10px;cursor:pointer}
.cg-psh:hover{filter:brightness(.98)}
.cg-pst{font-size:15px;font-weight:500;flex:1}
.cg-pstog{font-size:12px;color:var(--tx3);font-weight:500}
.cg-psb{display:none;padding:0}
.cg-psb.open{display:block}
.cg-pg{border-bottom:1px solid var(--bd);padding:14px 18px}
.cg-pg:last-child{border-bottom:none}
.cg-pgt{font-size:13px;font-weight:600;margin-bottom:10px;display:flex;align-items:center;gap:8px}
.cg-phs{font-size:11px;font-weight:500;padding:2px 8px;border-radius:10px;background:var(--bg3);color:var(--tx2);border:1px solid var(--bd2)}
.cg-ti{display:flex;align-items:flex-start;gap:8px;padding:6px 0;border-bottom:1px dashed var(--bd)}
.cg-ti:last-child{border-bottom:none}
.cg-cb{accent-color:var(--p2);width:15px;height:15px;margin-top:2px;flex-shrink:0;cursor:pointer}
.cg-tx{font-size:13px;color:var(--tx);line-height:1.5;flex:1}
.cg-tx.done{text-decoration:line-through;color:var(--tx3)}
.cg-tx input{width:100%;font-size:13px;border:1px solid var(--bd2);border-radius:6px;padding:4px 8px;font-family:inherit;background:var(--bg2);color:var(--tx)}
.cg-del{background:none;border:none;cursor:pointer;color:var(--tx3);font-size:15px;padding:0 2px;flex-shrink:0}
.cg-del:hover{color:var(--red)}
.cg-ar{display:flex;gap:8px;margin-top:10px}
.cg-ar input{flex:1;padding:7px 12px;font-size:13px;border:1px solid var(--bd2);border-radius:8px;font-family:inherit;background:var(--bg2);color:var(--tx)}
.cg-fr{margin-bottom:12px}
.cg-fr label{font-size:12px;font-weight:500;color:var(--tx2);display:block;margin-bottom:4px}
.cg-fr input,.cg-fr select,.cg-fr textarea{width:100%;padding:8px 12px;font-size:13px;border:1px solid var(--bd2);border-radius:8px;font-family:inherit;background:var(--bg2);color:var(--tx)}
.cg-fr textarea{min-height:70px;resize:vertical;line-height:1.6}
.cg-fa{display:flex;gap:8px;flex-wrap:wrap}
.cg-apf{border:1.5px solid var(--p2b);border-radius:var(--r);padding:16px;margin-top:8px;display:none;background:var(--bg)}
.cg-apf.open{display:block}
.cg-addb{width:100%;border:2px dashed var(--bd2);background:none;border-radius:8px;padding:10px;font-size:13px;color:var(--tx3);cursor:pointer;font-family:inherit;transition:all .15s;margin-top:8px}
.cg-addb:hover{border-color:var(--p2);color:var(--p2);background:var(--p2l)}
.cg-foot{border-top:1px solid var(--bd);padding-top:16px;margin-top:8px;font-size:12px;color:var(--tx3);line-height:1.8}
</style>`;
  }

  // ── EVENT BINDING ─────────────────────────────────────────────────────────
  private _bind(): void {
    const d = this.domElement;

    // Static header buttons
    this._on(d, '#cgSync',  'click', () => this._sync());
    this._on(d, '#cgEdit',  'click', () => this._toggleEdit());
    this._on(d, '#cgExp',   'click', () => this._export());

    // Tab buttons
    ['overview','agenda','tasks'].forEach(tab => {
      this._on(d, '#cgt-' + tab, 'click', () => this._showTab(tab));
    });

    // ── Delegated CLICK ──────────────────────────────────────────────────
    d.addEventListener('click', (e: Event) => {
      let el = e.target as HTMLElement;
      let depth = 0;
      while (el && depth < 6) {
        const a = el.getAttribute('data-a');
        if (a) { this._handleClick(a, el); return; }
        el = el.parentElement as HTMLElement;
        depth++;
      }
    });

    // ── Delegated CHANGE (checkboxes + selects) ──────────────────────────
    d.addEventListener('change', (e: Event) => {
      const el = e.target as HTMLElement;
      const a = el.getAttribute('data-a');
      if (!a) return;
      if (a === 'cb') {
        const cb = el as HTMLInputElement;
        this._toggleTask(+el.getAttribute('data-pi'), +el.getAttribute('data-pri'), +el.getAttribute('data-ti'), cb.checked);
      } else if (a === 'atag') {
        const sel = el as HTMLSelectElement;
        const i = +el.getAttribute('data-i');
        const lm: Record<string,string> = {all:'All',p1:'Pillar 1: Dynamics',p2:'Pillar 2: Automation',p3:'Pillar 3: Government',p4:'Pillar 4: Technical'};
        this._a[i].tag = sel.value;
        this._a[i].tagLabel = lm[sel.value] || sel.value;
        this._renderAgenda();
      }
    });

    // ── Delegated INPUT ──────────────────────────────────────────────────
    d.addEventListener('input', (e: Event) => {
      const el = e.target as HTMLInputElement;
      const a = el.getAttribute('data-a');
      if (!a) return;
      const i = +el.getAttribute('data-i');
      const pi = +el.getAttribute('data-pi');
      if (a === 'ttxt')   { this._p[pi].projects[+el.getAttribute('data-pri')].tasks[+el.getAttribute('data-ti')].text = el.value; }
      else if (a === 'ch'){ this._p[pi].champion = el.value; }
      else if (a === 'ds'){ this._p[pi].desc = el.value; }
      else if (a === 'an'){ this._a[i].num = el.value; }
      else if (a === 'at'){
        this._a[i].time = el.value;
        let _tot = 0;
        for (let _j = 0; _j < this._a.length; _j++) { const _m = parseInt(this._a[_j].time); if (!isNaN(_m)) _tot += _m; }
        const _amEl = this.domElement.querySelector('.cg-am');
        if (_amEl) _amEl.textContent = 'Total: ~' + _tot + ' minutes';
      }
      else if (a === 'atl'){ this._a[i].title = el.value; }
      else if (a === 'ap'){ this._a[i].points[+el.getAttribute('data-pi')] = el.value; }
      else if (a === 'ano'){ this._a[i].note = el.value; }
    });

    // ── Delegated FOCUSOUT (champion save on blur) ────────────────────────
    d.addEventListener('focusout', (e: Event) => {
      const el = e.target as HTMLElement;
      if (el.getAttribute('data-a') === 'ch') { this._saveChampion(+el.getAttribute('data-pi')); }
    });

    // ── Delegated KEYDOWN (Enter in new-task row) ─────────────────────────
    d.addEventListener('keydown', (e: KeyboardEvent) => {
      if (e.key !== 'Enter') return;
      const el = e.target as HTMLElement;
      if (el.getAttribute('data-a') === 'nti') {
        this._addTask(+el.getAttribute('data-pi'), +el.getAttribute('data-pri'), el.getAttribute('data-prid'));
      }
    });
  }

  private _on(root: HTMLElement, sel: string, evt: string, fn: EventListener): void {
    const el = root.querySelector(sel);
    if (el) el.addEventListener(evt, fn);
  }

  private _handleClick(a: string, el: HTMLElement): void {
    const pi  = +el.getAttribute('data-pi');
    const pri = +el.getAttribute('data-pri');
    const ti  = +el.getAttribute('data-ti');
    const i   = +el.getAttribute('data-i');
    switch (a) {
      case 'tpil':  this._togglePillar(pi); break;
      case 'dtsk':  this._delTask(pi, pri, ti); break;
      case 'atsk':  this._addTask(pi, pri, el.getAttribute('data-prid')); break;
      case 'dprj':  this._delProject(pi, pri); break;
      case 'tapf':  this._toggleAPF(el.getAttribute('data-id')); break;
      case 'aprj':  this._addProject(pi, el.getAttribute('data-plid')); break;
      case 'tag':   this._toggleAgenda(i); break;
      case 'dai':   this._delAgendaItem(i); break;
      case 'aai':   this._addAgendaItem(); break;
      case 'dap':   this._delAgendaPoint(i, +el.getAttribute('data-pi')); break;
      case 'aap':   this._addAgendaPoint(i); break;
    }
  }

  // ── TABS / EDIT / EXPORT ──────────────────────────────────────────────────
  private _showTab(t: string): void {
    this._activeTab = t;
    for (const id of ['overview','agenda','tasks']) {
      const v = this.domElement.querySelector('#cgv-' + id) as HTMLElement;
      const b = this.domElement.querySelector('#cgt-' + id);
      if (v) v.className = id !== t ? 'cg-hide' : '';
      if (b) b.className = 'cg-tab' + (id === t ? ' cg-ta' : '');
    }
    if (t === 'tasks')    this._renderTasks();
    if (t === 'overview') this._renderOverview();
    if (t === 'agenda')   this._renderAgenda();
  }

  private _toggleEdit(): void {
    this._editMode = !this._editMode;
    const lbl  = this.domElement.querySelector('#cgELbl');
    const pill = this.domElement.querySelector('#cgEpill');
    if (lbl)  lbl.textContent  = this._editMode ? '✎ Disable editing' : '✎ Enable editing';
    if (pill) pill.className   = 'cg-epill' + (this._editMode ? ' on' : '');
    this._renderAll();
  }

  private _export(): void {
    const state = JSON.stringify({pillars: this._p, agendaItems: this._a, exported: Date.now()});
    const tag   = '<script>window._embeddedState=' + state + ';<\/script>';
    fetch(this.context.pageContext.web.absoluteUrl + '/SiteAssets/carbon_group_committee.html', {credentials:'same-origin'})
      .then(r => r.text())
      .then(html => {
        const out = html.replace('</head>', tag + '</head>');
        const b   = new Blob([out], {type:'text/html'});
        const a   = document.createElement('a');
        a.href     = URL.createObjectURL(b);
        a.download = 'Carbon-Group-AI-Committee-' + new Date().toISOString().slice(0,10) + '.html';
        a.click();
      })
      .catch(() => {
        alert('Export requires carbon_group_committee.html in Site Assets. Download it from the standalone version instead.');
      });
  }

  // ── RENDER ────────────────────────────────────────────────────────────────
  private _renderAll(): void {
    this._renderOverview();
    this._renderAgenda();
    if (this._activeTab === 'tasks') this._renderTasks();
  }

  private _renderOverview(): void {
    const pg = this.domElement.querySelector('#cgv-overview');
    if (!pg) return;
    let html = this._introOverview() + this._renderProgress() + '<div class="cg-grid">';
    for (let pi = 0; pi < this._p.length; pi++) {
      const pl  = this._p[pi];
      const col = PC[pl.id] || PC.p1;
      let tasks: ITask[] = [];
      for (let j = 0; j < pl.projects.length; j++) tasks = tasks.concat(pl.projects[j].tasks);
      let done = 0; for (let k = 0; k < tasks.length; k++) { if (tasks[k].done) done++; }
      const champH = this._editMode
        ? `<input class="cg-chin" value="${this._e(pl.champion)}" data-a="ch" data-pi="${pi}" placeholder="Champion name...">`
        : `<span>${this._e(pl.champion)}</span>`;
      const descH = this._editMode
        ? `<div class="cg-pd"><textarea data-a="ds" data-pi="${pi}">${this._e(pl.desc)}</textarea></div>`
        : `<div class="cg-pd">${this._e(pl.desc)}</div>`;
      html += `<div class="cg-pc cg-${col.cls}">
        <div class="cg-ph">
          <div class="cg-pn">${col.num}</div>
          <div class="cg-pt">${this._e(pl.title)}</div>
          <div class="cg-pch">Champion: ${champH}</div>
        </div>
        <div class="cg-pb">${descH}
          <span class="cg-tc">${done}/${tasks.length} tasks complete</span>
        </div></div>`;
    }
    html += '</div>';
    pg.innerHTML = html;
  }

  private _renderProgress(): string {
    let html = '<div class="cg-pgw"><div class="cg-pgt">Overall progress by pillar</div>';
    for (let i = 0; i < this._p.length; i++) {
      const pl = this._p[i];
      let tasks: ITask[] = [];
      for (let j = 0; j < pl.projects.length; j++) tasks = tasks.concat(pl.projects[j].tasks);
      let done = 0; for (let k = 0; k < tasks.length; k++) { if (tasks[k].done) done++; }
      const pct = tasks.length ? Math.round(done / tasks.length * 100) : 0;
      const col = (PC[pl.id] || PC.p1).bar;
      html += `<div class="cg-pgr">
        <div class="cg-pgl" style="color:${col}">${this._e(pl.title)}</div>
        <div class="cg-pgb"><div class="cg-pgf" style="width:${pct}%;background:${col}"></div></div>
        <div class="cg-pgp">${done}/${tasks.length} (${pct}%)</div>
      </div>`;
    }
    return html + '</div>';
  }

  private _renderAgenda(): void {
    const c = this.domElement.querySelector('#cgv-agenda');
    if (!c) return;
    const tagOpts  = ['all','p1','p2','p3','p4'];
    const tagLabels: Record<string,string> = {all:'All',p1:'Pillar 1: Dynamics',p2:'Pillar 2: Automation',p3:'Pillar 3: Government',p4:'Pillar 4: Technical'};
    const tagMap:   Record<string,string>  = {all:'tg-all',p1:'tg-p1',p2:'tg-p2',p3:'tg-p3',p4:'tg-p4'};
    let total = 0;
    for (let i = 0; i < this._a.length; i++) { const m = parseInt(this._a[i].time); if (!isNaN(m)) total += m; }
    let html = this._introAgenda() +
      `<div class="cg-aw"><div class="cg-ah"><h2>Monthly AI Committee &mdash; Standard Agenda</h2><span class="cg-am">Total: ~${total} minutes</span></div>`;

    for (let i = 0; i < this._a.length; i++) {
      const item = this._a[i];
      const tCls = tagMap[item.tag] || 'tg-all';
      const numH = this._editMode
        ? `<input style="width:36px;font-size:13px;font-weight:600;border:1px solid var(--bd2);border-radius:5px;padding:2px 4px;font-family:inherit;background:var(--bg2)" value="${this._e(item.num)}" data-a="an" data-i="${i}">`
        : `<span class="cg-anum">${this._e(item.num)}</span>`;
      const tmH = this._editMode
        ? `<input style="width:64px;font-size:12px;border:1px solid var(--bd2);border-radius:5px;padding:2px 4px;font-family:inherit;background:var(--bg2)" value="${this._e(item.time)}" data-a="at" data-i="${i}">`
        : `<span class="cg-atm">${this._e(item.time)}</span>`;
      const tlH = this._editMode
        ? `<input style="flex:1;font-size:14px;font-weight:500;border:1px solid var(--bd2);border-radius:5px;padding:3px 7px;font-family:inherit;background:var(--bg2)" value="${this._e(item.title)}" data-a="atl" data-i="${i}">`
        : `<span class="cg-atl">${this._e(item.title)}</span>`;
      let tagH: string;
      if (this._editMode) {
        tagH = `<select style="font-size:11px;border:1px solid var(--bd2);border-radius:5px;padding:2px 5px;font-family:inherit;background:var(--bg2)" data-a="atag" data-i="${i}">`;
        for (let oi = 0; oi < tagOpts.length; oi++) {
          tagH += `<option value="${tagOpts[oi]}"${item.tag===tagOpts[oi]?' selected':''}>${tagLabels[tagOpts[oi]]}</option>`;
        }
        tagH += '</select>';
      } else {
        tagH = `<span class="cg-atg ${tCls}">${this._e(item.tagLabel)}</span>`;
      }
      const delBtn = this._editMode ? `<button class="cg-del" style="margin-left:4px" data-a="dai" data-i="${i}" title="Delete">&#10005;</button>` : '';
      let ptsH = '';
      if (this._editMode) {
        for (let pi = 0; pi < item.points.length; pi++) {
          ptsH += `<li style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
            <input style="flex:1;font-size:13px;border:1px solid var(--bd2);border-radius:5px;padding:4px 7px;font-family:inherit;background:var(--bg)" value="${this._e(item.points[pi])}" data-a="ap" data-i="${i}" data-pi="${pi}">
            <button class="cg-del" data-a="dap" data-i="${i}" data-pi="${pi}">&#10005;</button></li>`;
        }
        ptsH += `<li style="margin-top:4px"><button class="cg-addb" style="margin-top:0;padding:5px 10px;font-size:12px;width:auto" data-a="aap" data-i="${i}">+ Add point</button></li>`;
      } else {
        for (let pi = 0; pi < item.points.length; pi++) ptsH += `<li>${this._e(item.points[pi])}</li>`;
      }
      const noteH = this._editMode
        ? `<div style="margin-top:10px"><label style="font-size:11px;color:var(--tx3);font-weight:500">&#128204; Note</label><textarea style="width:100%;margin-top:4px;font-size:12px;border:1px solid var(--bd2);border-radius:6px;padding:6px 8px;font-family:inherit;background:var(--bg);color:var(--tx2);resize:vertical;min-height:48px" data-a="ano" data-i="${i}">${this._e(item.note)}</textarea></div>`
        : `<p class="cg-note">&#128204; ${this._e(item.note)}</p>`;
      const hdrAttr = this._editMode ? ' style="cursor:default;flex-wrap:wrap;gap:8px"' : ` data-a="tag" data-i="${i}"`;
      html += `<div class="cg-ai">
        <div class="cg-aih"${hdrAttr}>${numH}${tmH}${tlH}${tagH}
          ${this._editMode ? delBtn : `<span class="cg-atog" id="cgatog-${i}">&#9660; Detail</span>`}
        </div>
        <div class="cg-ab${this._editMode?' open':''}" id="cgab-${i}">
          <ul style="${this._editMode?'list-style:none;padding-left:0':'padding-left:18px'}">${ptsH}</ul>
          ${noteH}
        </div></div>`;
    }
    if (this._editMode) html += `<div style="padding:12px 18px"><button class="cg-addb" data-a="aai">+ Add agenda item</button></div>`;
    html += '</div>';
    c.innerHTML = html;
  }

  private _renderTasks(): void {
    const c = this.domElement.querySelector('#cgv-tasks');
    if (!c) return;
    let html = this._introTasks();
    for (let pi = 0; pi < this._p.length; pi++) {
      const pl  = this._p[pi];
      const col = PC[pl.id] || PC.p1;
      let all: ITask[] = [];
      for (let j = 0; j < pl.projects.length; j++) all = all.concat(pl.projects[j].tasks);
      let done = 0; for (let k = 0; k < all.length; k++) { if (all[k].done) done++; }
      let projs = '';
      for (let pri = 0; pri < pl.projects.length; pri++) {
        const pr = pl.projects[pri];
        let rows = '';
        for (let ti = 0; ti < pr.tasks.length; ti++) {
          const t = pr.tasks[ti];
          rows += `<div class="cg-ti">
            <input type="checkbox" class="cg-cb"${t.done?' checked':''} data-a="cb" data-pi="${pi}" data-pri="${pri}" data-ti="${ti}">
            <span class="cg-tx${t.done?' done':''}">
              ${this._editMode ? `<input value="${this._e(t.text)}" data-a="ttxt" data-pi="${pi}" data-pri="${pri}" data-ti="${ti}">` : this._e(t.text)}
            </span>
            ${this._editMode ? `<button class="cg-del" data-a="dtsk" data-pi="${pi}" data-pri="${pri}" data-ti="${ti}">&#10005;</button>` : ''}
          </div>`;
        }
        const addRow = this._editMode
          ? `<div class="cg-ar"><input type="text" data-a="nti" data-pi="${pi}" data-pri="${pri}" data-prid="${this._e(pr.id)}" placeholder="Add a new task...">
             <button class="cg-btn cg-bg cg-bsm" data-a="atsk" data-pi="${pi}" data-pri="${pri}" data-prid="${this._e(pr.id)}">Add</button></div>`
          : '';
        projs += `<div class="cg-pg">
          <div class="cg-pgt">${this._e(pr.title)}<span class="cg-phs">${this._e(pr.phase)}</span>
            ${this._editMode ? `<button class="cg-btn cg-bdanger cg-bsm" style="margin-left:auto" data-a="dprj" data-pi="${pi}" data-pri="${pri}">&#10005; Delete</button>` : ''}
          </div>${rows}${addRow}</div>`;
      }
      const apf = this._editMode
        ? `<div style="padding:12px 18px">
            <button class="cg-addb" data-a="tapf" data-id="cgapf-${pl.id}">+ Add new project to this pillar</button>
            <div class="cg-apf" id="cgapf-${pl.id}">
              <div style="font-size:14px;font-weight:500;margin-bottom:12px">New project</div>
              <div class="cg-fr"><label>Project title</label><input type="text" id="cgapft-${pl.id}" placeholder="Project title..."></div>
              <div class="cg-fr"><label>Phase / timeline</label><input type="text" id="cgapfp-${pl.id}" placeholder="e.g. Month 2 — Core Engine"></div>
              <div class="cg-fa">
                <button class="cg-btn cg-bg" data-a="aprj" data-pi="${pi}" data-plid="${pl.id}">Add project</button>
                <button class="cg-btn cg-bw" style="background:var(--bg3);color:var(--tx2);border-color:var(--bd2)" data-a="tapf" data-id="cgapf-${pl.id}">Cancel</button>
              </div></div></div>`
        : '';
      html += `<div class="cg-ps">
        <div class="cg-psh" data-a="tpil" data-pi="${pi}" style="background:${col.bar}15;border-bottom-color:${col.bar}40">
          <div class="cg-pst" style="color:${col.bar}">${this._e(pl.title)}</div>
          <span style="font-size:12px;color:${col.bar};margin-right:8px">${done}/${all.length} complete</span>
          <span class="cg-pstog" id="cgpstog-${pi}">&#9650; Collapse</span>
        </div>
        <div class="cg-psb open" id="cgpsb-${pi}">${projs}${apf}</div>
      </div>`;
    }
    c.innerHTML = html;
  }

  // ── TASK ACTIONS ──────────────────────────────────────────────────────────
  private _toggleTask(pi: number, pri: number, ti: number, checked: boolean): void {
    const task = this._p[pi].projects[pri].tasks[ti];
    task.done  = checked;
    this._renderTasks();
    this._refreshProgress();
    if (!this._spConnected || !task.spId) return;
    this._setStatus('cg-sng', '&#8593; Saving...');
    this._merge(task.spId, {Done: task.done, CompletedDate: task.done ? new Date().toISOString() : null})
      .then(() => this._saved(task.done ? 'Task completed' : 'Task reopened'))
      .catch(() => this._setStatus('cg-err', '&#10005; Save failed'));
  }

  private _saveChampion(pi: number): void {
    if (!this._spConnected) return;
    const pl = this._p[pi];
    const items: ITask[] = [];
    for (let j = 0; j < pl.projects.length; j++) {
      for (let k = 0; k < pl.projects[j].tasks.length; k++) {
        if (pl.projects[j].tasks[k].spId) items.push(pl.projects[j].tasks[k]);
      }
    }
    if (!items.length) return;
    this._setStatus('cg-sng', '&#8593; Saving champion...');
    let done = 0;
    for (let i = 0; i < items.length; i++) {
      this._merge(items[i].spId, {PillarChampion: pl.champion})
        .then(() => { done++; if (done === items.length) this._saved('Champion saved'); })
        .catch(() => { /* silent */ });
    }
  }

  private _addTask(pi: number, pri: number, prid: string): void {
    const inp = this.domElement.querySelector(`[data-a="nti"][data-prid="${prid}"]`) as HTMLInputElement;
    const txt = inp && inp.value && inp.value.trim();
    if (!txt) return;
    const pl = this._p[pi], pr = pl.projects[pri];
    const newTask: ITask = {id: this._uid(), text: txt, done: false};
    pr.tasks.push(newTask);
    if (inp) inp.value = '';
    this._renderTasks();
    this._refreshProgress();
    if (!this._spConnected) return;
    this._setStatus('cg-sng', '&#8593; Adding task...');
    this._create({
      Title: txt, Pillar: pl.title, PillarID: pl.id,
      PillarOrder: ['p1','p2','p3','p4'].indexOf(pl.id) + 1,
      PillarChampion: pl.champion, Project: pr.title, Phase: pr.phase,
      ProjectOrder: pri + 1, TaskOrder: pr.tasks.length, Done: false
    }).then((id: number) => { newTask.spId = id; this._saved('Task added'); })
      .catch(() => this._setStatus('cg-err', '&#10005; Add failed — saved locally'));
  }

  private _delTask(pi: number, pri: number, ti: number): void {
    if (!confirm('Remove this task?')) return;
    const task = this._p[pi].projects[pri].tasks[ti];
    this._p[pi].projects[pri].tasks.splice(ti, 1);
    this._renderTasks();
    this._refreshProgress();
    if (!this._spConnected || !task.spId) return;
    this._setStatus('cg-sng', '&#8593; Deleting...');
    this._del(task.spId)
      .then(() => this._saved('Task deleted'))
      .catch(() => this._setStatus('cg-err', '&#10005; Delete failed'));
  }

  private _delProject(pi: number, pri: number): void {
    if (!confirm(`Delete "${this._p[pi].projects[pri].title}" and all its tasks?`)) return;
    const tasks = this._p[pi].projects[pri].tasks.slice();
    this._p[pi].projects.splice(pri, 1);
    this._renderTasks();
    this._refreshProgress();
    if (!this._spConnected) return;
    const sp = tasks.filter(t => !!t.spId);
    if (!sp.length) { this._saved('Project deleted'); return; }
    this._setStatus('cg-sng', '&#8593; Deleting...');
    let done = 0;
    for (let i = 0; i < sp.length; i++) {
      this._del(sp[i].spId)
        .then(() => { done++; if (done === sp.length) this._saved('Project deleted'); })
        .catch(() => { /* silent */ });
    }
  }

  private _addProject(pi: number, plid: string): void {
    const tEl = this.domElement.querySelector(`#cgapft-${plid}`) as HTMLInputElement;
    const pEl = this.domElement.querySelector(`#cgapfp-${plid}`) as HTMLInputElement;
    const title = tEl && tEl.value && tEl.value.trim();
    if (!title) { alert('Please enter a project title'); return; }
    this._p[pi].projects.push({id: this._uid(), title, phase: pEl ? pEl.value.trim() : '', tasks: []});
    this._renderTasks();
    this._refreshProgress();
  }

  private _togglePillar(pi: number): void {
    const b = this.domElement.querySelector(`#cgpsb-${pi}`);
    const t = this.domElement.querySelector(`#cgpstog-${pi}`);
    if (!b || !t) return;
    const open = b.classList.toggle('open');
    t.textContent = open ? '▲ Collapse' : '▼ Expand';
  }

  private _toggleAPF(id: string): void {
    const el = this.domElement.querySelector(`#${id}`);
    if (el) el.classList.toggle('open');
  }

  // ── AGENDA ACTIONS ────────────────────────────────────────────────────────
  private _toggleAgenda(i: number): void {
    const b = this.domElement.querySelector(`#cgab-${i}`);
    const t = this.domElement.querySelector(`#cgatog-${i}`);
    if (!b || !t) return;
    const open = b.classList.toggle('open');
    t.textContent = open ? '▲ Hide' : '▼ Detail';
  }
  private _addAgendaItem(): void {
    const num = this._a.length + 1;
    const n = (num < 10 ? '0' : '') + num;
    this._a.push({num:n, time:'5 min', title:'New agenda item', tag:'all', tagLabel:'All', points:['Add your points here'], note:'Add a note here.'});
    this._renderAgenda();
  }
  private _delAgendaItem(i: number): void {
    if (!confirm('Remove this agenda item?')) return;
    this._a.splice(i, 1);
    this._renderAgenda();
  }
  private _delAgendaPoint(i: number, pi: number): void { this._a[i].points.splice(pi, 1); this._renderAgenda(); }
  private _addAgendaPoint(i: number): void              { this._a[i].points.push('New point'); this._renderAgenda(); }

  // ── SP CONNECTIVITY ───────────────────────────────────────────────────────
  private _connect(): void {
    this._setStatus('cg-cng', '&#9676; Connecting to list...');
    this._fetchList();
  }

  private _sync(): void {
    this._setStatus('cg-cng', '&#9676; Syncing...');
    this._fetchList();
  }

  private _fetchList(): void {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${SP_LIST}')/items` +
      `?$select=Id,Title,Pillar,PillarID,PillarOrder,PillarChampion,Project,Phase,ProjectOrder,TaskOrder,Done` +
      `&$orderby=PillarOrder,ProjectOrder,TaskOrder&$top=500`;

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => {
        if (!r.ok) { throw new Error('HTTP ' + r.status); }
        return r.json();
      })
      .then((data: {value: any[]}) => {
        const items = data && data.value;
        if (!items || !items.length) {
          this._setStatus('cg-err', '&#10005; List empty — run Seed-Lists.ps1');
          return;
        }
        this._p = this._buildFromItems(items);
        this._spConnected = true;
        this._setStatus('cg-con', `&#9679; Connected &middot; AI_Committee &middot; ${items.length} tasks &middot; ${new Date().toLocaleTimeString()}`);
        this._renderAll();
      })
      .catch((err: Error) => {
        this._spConnected = false;
        if (err && err.message && err.message.indexOf('404') !== -1) {
          this._setStatus('cg-err', '&#10005; List not found — run Seed-Lists.ps1 first');
        } else {
          this._setStatus('cg-err', '&#10005; Could not connect — check list name is AI_Committee');
        }
      });
  }

  private _buildFromItems(items: any[]): IPillar[] {
    const pMap: Record<string, IPillar>    = {};
    const prMap: Record<string, IProject>  = {};
    const order = ['p1','p2','p3','p4'];
    const defaults = this._defaultPillars();

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const pid  = item.PillarID || 'p1';
      if (!pMap[pid]) {
        pMap[pid] = {id: pid, title: item.Pillar || pid, champion: item.PillarChampion || 'TBD', desc: '', projects: []};
      }
      if (item.PillarChampion && item.PillarChampion !== 'TBD') pMap[pid].champion = item.PillarChampion;
      const prKey = pid + '||' + item.Project;
      if (!prMap[prKey]) {
        const proj: IProject = {id: 'pr_' + item.ProjectOrder, title: item.Project, phase: item.Phase || '', tasks: []};
        prMap[prKey] = proj;
        pMap[pid].projects.push(proj);
      }
      prMap[prKey].tasks.push({id: 'sp_' + item.Id, spId: item.Id, text: item.Title, done: !!item.Done});
    }

    // Restore pillar descriptions from defaults
    for (let di = 0; di < defaults.length; di++) {
      if (pMap[defaults[di].id]) pMap[defaults[di].id].desc = defaults[di].desc;
    }

    const result: IPillar[] = [];
    for (let oi = 0; oi < order.length; oi++) {
      if (pMap[order[oi]]) result.push(pMap[order[oi]]);
    }
    return result;
  }

  // ── SP REST HELPERS ───────────────────────────────────────────────────────
  private _merge(id: number, data: Record<string,any>): Promise<void> {
    const url  = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${SP_LIST}')/items(${id})`;
    const opts: ISPHttpClientOptions = {
      headers: {'Content-Type':'application/json;odata=nometadata','X-HTTP-Method':'MERGE','IF-MATCH':'*'},
      body: JSON.stringify(data)
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, opts)
      .then((r: SPHttpClientResponse) => { if (!r.ok && r.status !== 204) throw new Error('HTTP ' + r.status); });
  }

  private _create(data: Record<string,any>): Promise<number> {
    const url  = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${SP_LIST}')/items`;
    const opts: ISPHttpClientOptions = {
      headers: {'Content-Type':'application/json;odata=nometadata','Accept':'application/json;odata=nometadata'},
      body: JSON.stringify(data)
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, opts)
      .then((r: SPHttpClientResponse) => r.json())
      .then((d: {Id: number}) => d.Id);
  }

  private _del(id: number): Promise<void> {
    const url  = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${SP_LIST}')/items(${id})`;
    const opts: ISPHttpClientOptions = {
      headers: {'Content-Type':'application/json;odata=nometadata','X-HTTP-Method':'DELETE','IF-MATCH':'*'},
      body: ''
    };
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, opts)
      .then((r: SPHttpClientResponse) => { if (!r.ok && r.status !== 204) throw new Error('HTTP ' + r.status); });
  }

  // ── STATUS ────────────────────────────────────────────────────────────────
  private _setStatus(cls: string, msg: string): void {
    const el = this.domElement.querySelector('#cgSt') as HTMLElement;
    if (!el) return;
    el.className = 'cg-st ' + cls;
    el.innerHTML = msg;
  }

  private _saved(msg: string): void {
    this._setStatus('cg-sav', '&#10003; ' + msg);
    clearTimeout(this._stTimer);
    this._stTimer = setTimeout(() => {
      this._setStatus('cg-con', '&#9679; Connected &middot; AI_Committee');
    }, 2500);
  }

  private _refreshProgress(): void {
    const pg = this.domElement.querySelector('#cgv-overview');
    if (!pg || this._activeTab !== 'overview') return;
    const pgw = pg.querySelector('.cg-pgw');
    if (pgw) pgw.outerHTML = this._renderProgress();
  }

  // ── INTRO HTML ────────────────────────────────────────────────────────────
  private _introOverview(): string {
    return `<div class="cg-ibox">
      <p><strong>Purpose:</strong> This committee exists to coordinate Carbon Group's AI transformation across four structured pillars. Each pillar has a champion who owns delivery, reports monthly to the group, and works with Nathan between meetings.</p>
      <p><strong>Meeting cadence:</strong> Monthly. Each champion presents their pillar's progress, blockers, and next steps.</p>
      <p><strong>How it works:</strong> All AI projects from the roadmap have been re-mapped into the four pillars below. Champions own their pillar — not Nathan.</p>
    </div>`;
  }
  private _introAgenda(): string {
    return `<div class="cg-ibox"><p><strong>Standard monthly agenda</strong> — each meeting runs approximately 90 minutes. Nathan opens and closes every meeting.</p></div>`;
  }
  private _introTasks(): string {
    return `<div class="cg-ibox"><p><strong>All tasks from the roadmap, re-mapped into the four pillars.</strong> Champions use this view to track progress between meetings. Tick tasks as they are completed.</p></div>`;
  }

  // ── HELPERS ───────────────────────────────────────────────────────────────
  private _e(s: any): string {
    if (s === null || s === undefined) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
  private _uid(): string { return 't' + Date.now() + Math.random().toString(36).slice(2,5); }

  // ── DEFAULT DATA ──────────────────────────────────────────────────────────
  private _defaultPillars(): IPillar[] {
    return [
      {id:'p1',title:'Dynamics CRM',champion:'TBD',
       desc:'Build and configure Microsoft Dynamics as the single source of truth for all client data, job records, and communications. Everything else in the transformation depends on this foundation being correct.',
       projects:[
        {id:'pr1',title:'Dynamics CRM Build & Configuration',phase:'Month 1 — Foundation',tasks:[
          {id:'t1',text:'Set up Dynamics sandbox environment',done:false},
          {id:'t2',text:'Define and build the client record schema (entities, services, contacts, preferred window, goals)',done:false},
          {id:'t3',text:'Map all entity types: company, trust, individual, SMSF — and configure accordingly',done:false},
          {id:'t4',text:'Connect SharePoint/OneDrive as the document layer',done:false},
          {id:'t5',text:'Define job record fields: job type, status, assigned staff, queries, deadlines, completion state',done:false},
          {id:'t6',text:'Build the onboarding intake form (Content Snare equivalent — native to Dynamics)',done:false},
          {id:'t7',text:'Test data flow: onboarding to client record to job record',done:false},
          {id:'t8',text:'Define what job ready to start means as a system state',done:false}
        ]},
        {id:'pr2',title:'Client Data Audit & Migration',phase:'Month 1 — Foundation',tasks:[
          {id:'t9',text:'Audit existing client records across all firms — what exists, where it lives, how clean it is',done:false},
          {id:'t10',text:'Identify which clients have entity structure, preferred contact, and prior year documents accessible',done:false},
          {id:'t11',text:'Define minimum data required per client for the Pre-Start Engine to function',done:false},
          {id:'t12',text:'Build migration plan to move existing client data into Dynamics cleanly',done:false},
          {id:'t13',text:'Identify clients who need to be contacted to fill in missing information',done:false},
          {id:'t14',text:'Plan rollout sequence: which firms go first and in what order',done:false}
        ]},
        {id:'pr3',title:'Healthy, Wealthy & Wise — Dynamics Build',phase:'Month 3 — Advisory Layer',tasks:[
          {id:'t15',text:'Build the Healthy, Wealthy & Wise intake form in Dynamics (digital, shareable with client)',done:false},
          {id:'t16',text:'Configure separate completion for client and spouse/partner',done:false},
          {id:'t17',text:'Build the goal tracking view: business, personal, and family goals — visible to the advisor',done:false},
          {id:'t18',text:'Build the timeline view: 1yr, 3yr, 5yr, 10yr goals side by side',done:false},
          {id:'t19',text:'Configure the shared board between advisor and client: both can see and update before meetings',done:false},
          {id:'t20',text:'Build the action tracking system: actions from each meeting, assigned to client or advisor with due dates',done:false},
          {id:'t21',text:'Make action status visible to both parties (like Planner in Teams)',done:false},
          {id:'t22',text:'Build the meeting history log: what was discussed, decided, and what is outstanding',done:false},
          {id:'t23',text:'Configure advisor prompts: system suggests follow-up points between meetings based on open actions',done:false},
          {id:'t24',text:'Build cross-referral tracking and opportunity flag system',done:false},
          {id:'t25',text:'Build the referral dashboard for partners: who has been referred, to which division, what was the result',done:false}
        ]}
      ]},
      {id:'p2',title:'Process & System Automation',champion:'TBD',
       desc:'Automate the end-to-end production workflow — from job scheduling and document collection through to client signing and lodgement. This pillar directly addresses the capacity constraint and eliminates the manual coordination burden on client managers.',
       projects:[
        {id:'pr4',title:'Job Scheduling Engine',phase:'Month 2 — Core Engine',tasks:[
          {id:'t26',text:'Build the preferred completion window field into client onboarding',done:false},
          {id:'t27',text:'Build the job scheduling logic: generate the next job instance based on services and window',done:false},
          {id:'t28',text:'Configure the trigger: 30 days before window opens, automatically initiate the Pre-Start process',done:false},
          {id:'t29',text:'Build the production calendar view: upcoming jobs, status, and assigned team members',done:false},
          {id:'t30',text:'Configure alerts: notify pod leaders when a job has not been initiated within the trigger window',done:false},
          {id:'t31',text:'Test scheduling logic across all job types and entity combinations',done:false},
          {id:'t32',text:'For existing clients: backfill preferred window data and generate first scheduled jobs',done:false}
        ]},
        {id:'pr5',title:'AI-Generated Personalised Client Checklists',phase:'Month 2 — Core Engine',tasks:[
          {id:'t33',text:'Build the checklist generation logic: map job type and entity type to required documents',done:false},
          {id:'t34',text:'Integrate prior year data: pull what was requested and received last year',done:false},
          {id:'t35',text:'Build the smart additions layer: flag known issues from prior year',done:false},
          {id:'t36',text:'Design the client-facing checklist format — clear, simple, itemised',done:false},
          {id:'t37',text:'Build the automated email/notification with deadline and window commitment language',done:false},
          {id:'t38',text:'Test across: individual, company, trust, and SMSF combinations',done:false},
          {id:'t39',text:'Build the response tracking system: log what has been received and what is outstanding',done:false}
        ]},
        {id:'pr6',title:'Client Query & Document Collection System',phase:'Month 2 — Core Engine',tasks:[
          {id:'t40',text:'Build the query tracking system in Dynamics: each query is a record with status',done:false},
          {id:'t41',text:'Build the document intake portal: clients upload via a link into SharePoint (not email)',done:false},
          {id:'t42',text:'Configure automatic document tagging: linked to the relevant job and query',done:false},
          {id:'t43',text:'Build the partial response logic: only resend unanswered queries',done:false},
          {id:'t44',text:'Configure automatic reminders if no response after X days',done:false},
          {id:'t45',text:'Build the job status update: all queries resolved means ready to start',done:false},
          {id:'t46',text:'Build team notification: alert accountant when job moves to ready to start',done:false},
          {id:'t47',text:'Ensure all responses are visible to team — not locked in one inbox',done:false},
          {id:'t48',text:'Test full loop: send, respond, partial, follow up, complete, job unlocks',done:false}
        ]},
        {id:'pr7',title:'Pre-Start Validation Check',phase:'Month 2 — Core Engine',tasks:[
          {id:'t49',text:'Build the pre-start validation logic: check all required documents against the checklist',done:false},
          {id:'t50',text:'Cross-reference received documents against prior year requirements',done:false},
          {id:'t51',text:'Flag any documents that appear missing or incomplete',done:false},
          {id:'t52',text:'Generate a pre-start summary for the accountant: what we have, what is confirmed, any flags',done:false},
          {id:'t53',text:'Block job from entering production queue until pre-start check is passed',done:false},
          {id:'t54',text:'Build the override: allow pod leader to manually approve a job if there is a valid reason',done:false}
        ]},
        {id:'pr8',title:'Document Compilation, Signing & Client Output',phase:'Month 2-3 — Quality Layer',tasks:[
          {id:'t55',text:'Build document compilation logic: bundle all job documents in correct order by entity type',done:false},
          {id:'t56',text:'Build cover letter generator: summarise tax outcomes, refunds/payables, key points for the year',done:false},
          {id:'t57',text:'Build the proactive insights section: flag payroll tax proximity, FBT, Div 7A, and other common issues',done:false},
          {id:'t58',text:'Integrate signing infrastructure into our own stack (not a third-party cost)',done:false},
          {id:'t59',text:'Configure multi-signature routing: different signatories for different entities in the same group',done:false},
          {id:'t60',text:'Build the signing reminder system: automated reminders until all documents are signed',done:false},
          {id:'t61',text:'Track signing status in Dynamics: visible to the full team',done:false},
          {id:'t62',text:'Notify client manager when signing is complete and lodgement is ready',done:false}
        ]},
        {id:'pr9',title:'Quarterly Meeting & Advisory System',phase:'Month 3 — Advisory Layer',tasks:[
          {id:'t63',text:'Build the quarterly meeting scheduler: automatically suggest meeting dates based on client preference',done:false},
          {id:'t64',text:'Build the advisor prep dashboard: surfaces client goals, open actions, recent job outcomes, and flags',done:false},
          {id:'t65',text:'Build the shared agenda tool: client and advisor can both add agenda items before the meeting',done:false},
          {id:'t66',text:'Build the post-meeting summary: advisor logs key points, new actions are created and assigned',done:false},
          {id:'t67',text:'Configure the follow-up prompt: advisor reminded to follow up on client actions between meetings',done:false},
          {id:'t68',text:'Train all partners and pod leaders on the advisory model and how to run the meetings',done:false}
        ]}
      ]},
      {id:'p3',title:'Integration with Government Agencies',champion:'TBD',
       desc:'Build and maintain the connections between our systems and external government bodies — primarily the ATO — to enable automated lodgement, real-time status updates, and compliance monitoring.',
       projects:[
        {id:'pr10',title:'ATO Lodgement Automation',phase:'Month 2-3 — Quality Layer',tasks:[
          {id:'t69',text:'Map all lodgement types: tax returns (individual, company, trust, SMSF), BAS, payroll, ASIC',done:false},
          {id:'t70',text:'Define the trigger: once all signatures received, automatically prepare lodgement package',done:false},
          {id:'t71',text:'Build the lodgement trigger in Dynamics: connects to ATO portal or intermediary',done:false},
          {id:'t72',text:'Configure lodgement status tracking: pending, submitted, accepted, rejected — visible in Dynamics',done:false},
          {id:'t73',text:'Build the error handling flow: if ATO rejects, flag to the relevant accountant with reason',done:false},
          {id:'t74',text:'Build the lodgement confirmation notification: notify client and update job status to complete',done:false}
        ]},
        {id:'pr11',title:'ATO Compliance Monitoring & Debt Alerts',phase:'Month 3 — Advisory Layer',tasks:[
          {id:'t75',text:'Build integration to pull ATO account balances and debt status per client',done:false},
          {id:'t76',text:'Configure alerts: notify client manager when a client has an outstanding ATO debt or payment plan',done:false},
          {id:'t77',text:'Surface ATO correspondence and notices into the client record in Dynamics',done:false},
          {id:'t78',text:'Build the proactive client notification: alert clients to upcoming payment obligations before they are due',done:false}
        ]},
        {id:'pr12',title:'ASIC & Other Agency Integrations',phase:'Month 3 — Advisory Layer',tasks:[
          {id:'t79',text:'Map all ASIC obligations per client: annual reviews, changes of officeholders, registered addresses',done:false},
          {id:'t80',text:'Build automated ASIC annual review reminder and lodgement workflow',done:false},
          {id:'t81',text:'Identify any other agency integrations required (state revenue offices, payroll tax portals)',done:false},
          {id:'t82',text:'Build the compliance calendar: shows all upcoming lodgement deadlines across all agencies for all clients',done:false}
        ]}
      ]},
      {id:'p4',title:'Technical Know-How',champion:'TBD',
       desc:'Build the AI layer that assists with the actual production of accounting and bookkeeping work — reviewing work for accuracy, checking for the best tax outcome, and eventually drafting components of standard compliance work.',
       projects:[
        {id:'pr13',title:'Workflow Documentation — All Core Job Types',phase:'Month 1 — Foundation',tasks:[
          {id:'t83',text:'Document the full workflow for: Individual Tax Return',done:false},
          {id:'t84',text:'Document the full workflow for: Company Tax Return + Financials',done:false},
          {id:'t85',text:'Document the full workflow for: Trust Tax Return + Financials',done:false},
          {id:'t86',text:'Document the full workflow for: BAS Lodgement',done:false},
          {id:'t87',text:'Document the full workflow for: SMSF Annual Return',done:false},
          {id:'t88',text:'Document the full workflow for: Payroll processing',done:false},
          {id:'t89',text:'For each job type: list every document required by entity type',done:false},
          {id:'t90',text:'For each job type: define what complete information looks like before production starts',done:false},
          {id:'t91',text:'For each job type: document the full review checklist (what reviewers currently check)',done:false},
          {id:'t92',text:'Identify variations: what changes based on client complexity or special circumstances',done:false}
        ]},
        {id:'pr14',title:'AI First-Level Review System',phase:'Month 2-3 — Quality Layer',tasks:[
          {id:'t93',text:'Define the first-level AI review checklist for each job type',done:false},
          {id:'t94',text:'For tax returns: check figures reconcile, prior year comparisons within range, all entities included',done:false},
          {id:'t95',text:'For financials: check balance sheet balances, P&L is complete, comparatives are correct',done:false},
          {id:'t96',text:'For BAS: check GST figures reconcile to coding, period is correct, lodgement date noted',done:false},
          {id:'t97',text:'Build the AI review layer: runs automatically when accountant marks job as complete',done:false},
          {id:'t98',text:'Generate an AI review report: list of checks passed, any flags for human attention',done:false},
          {id:'t99',text:'Build the workflow gate: job cannot go to human review until AI review is complete',done:false},
          {id:'t100',text:'For tax compliance jobs: compare final outcome to tax planning engagement — flag if more than 5% variance',done:false},
          {id:'t101',text:'Test AI review across all job types and validate accuracy with a senior accountant',done:false}
        ]},
        {id:'pr15',title:'Human Review Workflow (Technical + Commercial)',phase:'Month 2-3 — Quality Layer',tasks:[
          {id:'t102',text:'Define the technical review checklist: accuracy, completeness, compliance',done:false},
          {id:'t103',text:'Define the commercial review checklist: best tax position, missed opportunities, client flags',done:false},
          {id:'t104',text:'Build the two-stage review workflow in Dynamics: technical, commercial, approved',done:false},
          {id:'t105',text:'Configure routing: simple jobs go to client manager, complex jobs escalate to pod leader or partner',done:false},
          {id:'t106',text:'Build the feedback loop: reviewer sends job back with specific comments, not just a general rejection',done:false},
          {id:'t107',text:'Track review turnaround time: flag jobs sitting in review queue for more than X days',done:false},
          {id:'t108',text:'Build the approval gate: job cannot be compiled for client until both reviews are complete',done:false}
        ]},
        {id:'pr16',title:'AI-Assisted Meeting Preparation & Cross-Referral Prompts',phase:'Month 3 — Advisory Layer',tasks:[
          {id:'t109',text:'Configure AI-generated meeting prompts: based on Healthy, Wealthy & Wise data, suggest relevant topics',done:false},
          {id:'t110',text:'Build the cross-referral prompt layer: surface opportunities to refer into other divisions based on client goals',done:false},
          {id:'t111',text:'Build the gap analysis: which clients have only one service? Flag as expansion opportunities',done:false},
          {id:'t112',text:'Build the revenue attribution model: track cross-referral revenue back to the originating advisor',done:false}
        ]}
      ]}
    ];
  }

  private _defaultAgenda(): IAgenda[] {
    return [
      {num:'01',time:'5 min',title:'Welcome & Housekeeping',tag:'all',tagLabel:'All',
       points:['Nathan opens the meeting','Confirm attendees and apologies','Any urgent items to add to the agenda'],
       note:'Nathan chairs this section.'},
      {num:'02',time:'5 min',title:'Previous Actions Review',tag:'all',tagLabel:'All',
       points:['Review actions from last meeting','Confirm which are complete, in progress, or need to be carried forward','Any blockers to call out immediately'],
       note:'Champions confirm their own actions. Keep this tight — detail goes in pillar sections.'},
      {num:'03',time:'15 min',title:'Pillar 1: Dynamics CRM — Champion Presentation',tag:'p1',tagLabel:'Pillar 1: Dynamics',
       points:['Progress update since last meeting (tasks completed, milestones hit)','Current blockers or decisions needed from the group','Work planned before next meeting','Any requests for Nathan or others to assist with'],
       note:'Champion presents. Nathan and group ask questions. Agree on any actions before moving on.'},
      {num:'04',time:'15 min',title:'Pillar 2: Process & System Automation — Champion Presentation',tag:'p2',tagLabel:'Pillar 2: Automation',
       points:['Progress update since last meeting','Current blockers or decisions needed','Work planned before next meeting','Dependencies on other pillars (especially Dynamics)'],
       note:'Champion presents. Flag any cross-pillar dependencies clearly.'},
      {num:'05',time:'15 min',title:'Pillar 3: Government Agency Integration — Champion Presentation',tag:'p3',tagLabel:'Pillar 3: Government',
       points:['Progress update since last meeting','ATO and ASIC integration status','Any regulatory or compliance considerations to flag','Work planned before next meeting'],
       note:'Champion presents. Note that this pillar cannot fully proceed until Pillar 1 (Dynamics) has the job completion workflow in place.'},
      {num:'06',time:'15 min',title:'Pillar 4: Technical Know-How — Champion Presentation',tag:'p4',tagLabel:'Pillar 4: Technical',
       points:['Progress update since last meeting','Workflow documentation status','AI review build progress and accuracy results','Work planned before next meeting'],
       note:'Champion presents. This pillar feeds directly into Pillar 2 — document review must precede automation.'},
      {num:'07',time:'10 min',title:'Cross-Pillar Issues & Decisions',tag:'all',tagLabel:'All',
       points:['Any issues that span multiple pillars','Decisions that require the full group','Resourcing or prioritisation changes','Escalations that need Nathan to decide'],
       note:'Nathan facilitates. This is not a general discussion — it is for decisions only.'},
      {num:'08',time:'5 min',title:'Actions & Next Meeting',tag:'all',tagLabel:'All',
       points:['Confirm all actions: who, what, by when','Confirm date of next monthly meeting','Any between-meeting working sessions to schedule (optional, not compulsory)','Nathan closes the meeting'],
       note:'Keep this section tight. Every action must have a named owner and a date.'}
    ];
  }
}
