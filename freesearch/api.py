"""Zero-dependency dev server + adapter notes.

This wraps the framework-agnostic controller in Python's stdlib
`http.server` so the endpoint is runnable *today* with no pip install — handy
for local front-end integration and for proving the contract. It is NOT the
production server; for production, drop `handle_free_search` /
`handle_jurisdictions` into whatever the tenant already runs:

    FastAPI:
        @app.post('/free-search')
        async def free_search(req: Request):
            return handle_free_search(await req.json(), get_temmy_client())

    Flask:
        @app.post('/free-search')
        def free_search():
            return jsonify(handle_free_search(request.get_json(), client))

    Supabase Edge (Deno) / Lambda:
        call an HTTP shim in front of the same controller.

CORS is wide-open here for embedding; production must restrict
Access-Control-Allow-Origin to the tenant allow-list (TMH, Temmy portal,
approved introducer domains) and put IP rate-limiting in front.
"""
from __future__ import annotations

import json
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from urllib.parse import parse_qs, urlparse

from .controller import (handle_enrich, handle_free_search, handle_jurisdictions,
                          handle_lookup, handle_read_website, handle_suggest_classes)


def _make_client():
    """Build a TemmyClient from env. Imported lazily so importing this module
    never requires Temmy credentials (tests inject a fake client instead)."""
    import sys
    from pathlib import Path
    deploy = Path(__file__).resolve().parents[1] / 'deploy-v2-hotfix'
    sys.path.insert(0, str(deploy))
    from freesearch.temmy_pooled import PooledTemmyClient
    # .strip() defends against the trailing-newline / stray-space corruption
    # seen in the Temmy credential files (a '\n' makes an invalid HTTP header
    # value — exactly the failure the MCP Query Runs call surfaced).
    key = os.environ.get('TEMMY_API_KEY', '').strip()
    base = (os.environ.get('TEMMY_API_BASE_URL') or '').strip() or None
    # rate_limit_sec=0: TemmyDB is our own, free and unlimited — the client's
    # default 0.25s inter-call throttle only adds latency to the free search,
    # which the service parallelises across a thread pool. PooledTemmyClient
    # keeps one keep-alive Session so those parallel fetches reuse connections.
    kwargs = {'api_key': key, 'rate_limit_sec': 0}
    if base:
        kwargs['base_url'] = base
    return PooledTemmyClient(**kwargs)


# The same host serves the widget (decision 20 Jul: one deploy, one domain —
# partners embed a script tag; every widget update is one push, no per-partner
# copies to drift). Files come from freesearch/web/.
_WEB = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'web')
_PAGES = {
    '': 'free-search.html',                    # GET /  -> the wizard
    '/free-search': 'free-search.html',
    # The short wizard (Name -> Classes -> Results). Same file — the page
    # switches itself into quick mode from the pathname.
    '/uk-trademark-quick-search': 'free-search.html',
    # Standalone Class Builder (Jonathan, 21 Aug): Name -> Classes ->
    # Review & Send. Same file again — builder mode from the pathname. No
    # register search: the class tools feed a review screen where classes,
    # descriptions and UKIPO terms are amended, then SENT (to a Deal's
    # Classes & Terms when opened from Zoho, or emailed as a CSV to a new
    # Class Tools lead from the public site).
    '/class-builder': 'free-search.html',
    # Full search report at a unique URL (Jonathan, 21 Aug): the emailed
    # Free/Quick Search reports link here; the page renders the COMPLETE
    # stored result (all flagged marks) in the Sector-Report design and
    # offers a print-to-PDF download. /report/<session> also works — see
    # the startswith route in do_GET.
    '/search-report': 'report.html',
    '/class-assistant': 'class-assistant.html',
    '/search-bar': 'search-bar.html',
    '/search-box': 'search-box.html',          # compact drop-anywhere entry point
    # free-search.html's CONFIG.BRAND_AUDIT_URL points at '/brand-audit/', so
    # without this the "Request a Brand Audit" button on the results screen
    # 404s. The file existed and was built; it was simply never routed.
    '/brand-audit': 'brand-audit.html',
    # Client-facing audit checkout (Jonathan, 27 Aug): the Version 1 flow
    # that FOLLOWS Free/Quick Search — 9 pages from contact capture to
    # thank-you + fact find, pre-populated from the journey session (?s=).
    # The multi-brand tool above stays parked and untouched.
    '/audit': 'audit-checkout.html',
    # Standalone staff tool over POST /suggest-classes (Jonathan, 10 Aug).
    # Same endpoint Free Search and Brand Audit call — one agent, three UIs.
    '/class-agent': 'class-agent.html',
}

# Partners drop this one line into their page:
#   <script src="https://<host>/embed.js" data-tenant="acme" async></script>
# It injects the wizard in an iframe pinned to this host, so the widget is
# always the deployed version and the tenant id rides along.
#
# data-widget picks which one: 'free-search' (default, the full wizard),
# 'search-box' (the compact entry point), 'class-assistant', 'search-bar'.
# 'search-box' additionally accepts data-target (where the wizard lives) and
# data-journey (the journey function's base URL, since the engine and the
# Edge Function are different hosts in production).
#
# min-height is per-widget: the wizard needs room, the compact box would look
# absurd in a 640px frame. Both then self-report their real height.
_EMBED_JS = """(function(){
  var s=document.currentScript, t=(s&&s.dataset.tenant)||'tmh';
  var page=(s&&s.dataset.widget)||'free-search';
  var q='?tenant='+encodeURIComponent(t)+'&embed=1';
  if(page==='search-box'){
    if(s.dataset.variant) q+='&variant='+encodeURIComponent(s.dataset.variant);
    if(s.dataset.style)   q+='&style='+encodeURIComponent(s.dataset.style);
    if(s.dataset.target)  q+='&target='+encodeURIComponent(s.dataset.target);
    if(s.dataset.journey) q+='&journey='+encodeURIComponent(s.dataset.journey);
  }
  // Domain-hosted wizard pages (20 Aug): the search box hands off to a
  // WordPress page carrying ?s=<session>&screen=&q=&journey= — forward those
  // into the iframe so the embedded wizard adopts the session instead of
  // starting over. Only the known resume keys pass through.
  try{
    var hp=new URLSearchParams(window.location.search);
    // 'fs' is the WordPress-safe alias for the session param: WP reserves
    // ?s= for site search and 404s page URLs carrying it (21 Aug). The
    // host page uses fs; the iframe wizard still receives s.
    ['s','fs','screen','q','journey','searchbase','deal'].forEach(function(k){
      var v=hp.get(k); if(v) q+='&'+(k==='fs'?'s':k)+'='+encodeURIComponent(v);
    });
  }catch(e){}
  var minH = page==='search-box' ? (s.dataset.style==='bar'?'70px':'200px') : '760px';
  // Change spec 22 Aug (section 3): reserve space immediately so the WP page
  // never looks empty and the footer never jumps; show a branded loading
  // panel until the app's first height report proves it has rendered; fall
  // back to a static retry message on timeout.
  var wrap=document.createElement('div');
  wrap.style.cssText='position:relative;width:100%;min-height:'+minH;
  var f=document.createElement('iframe');
  f.src=s.src.replace(/embed\\.js.*$/, page+q);
  f.style.cssText='width:100%;min-height:'+minH+';border:0;display:block';
  f.setAttribute('title','Free Trademark Search');
  f.setAttribute('loading','lazy');
  wrap.appendChild(f);
  var panel=null, rot=null, tm=null, ready=false;
  if(page!=='search-box'){
    panel=document.createElement('div');
    panel.style.cssText='position:absolute;top:0;left:0;right:0;bottom:0;display:flex;'+
      'flex-direction:column;align-items:center;justify-content:center;background:#f7f8fa;'+
      'border:1px solid #e3e8ee;border-radius:12px;text-align:center;padding:28px;z-index:2;'+
      'font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif';
    panel.innerHTML='<div style="width:46px;height:46px;background:#E51652;color:#fff;'+
      'font-weight:800;font-size:18px;display:flex;align-items:center;justify-content:center;'+
      'transform:skewX(-12deg);margin:0 0 16px"><span style="transform:skewX(12deg)">TH</span></div>'+
      '<div style="font-weight:700;color:#1c2d3b;font-size:17px;margin-bottom:8px">'+
      'Loading your UK trademark search\\u2026</div>'+
      '<div data-bmsg style="color:#5b6b7a;font-size:14px;max-width:480px;line-height:1.55"></div>';
    wrap.appendChild(panel);
    var msgs=[
      'A register search helps identify earlier marks that may affect a new application.',
      'Challenges can arise during examination, opposition or after registration, so early research matters.',
      'A word search is an important first step \\u2014 logos, prior use and activity outside the register can also affect risk.'
    ];
    var mi=0, msgEl=panel.querySelector('[data-bmsg]');
    msgEl.textContent=msgs[0];
    rot=setInterval(function(){ mi=(mi+1)%msgs.length; msgEl.textContent=msgs[mi]; },5000);
    var arm=function(){
      tm=setTimeout(function(){
        if(ready) return;
        clearInterval(rot);
        msgEl.innerHTML='The search tool is taking longer than usual to load. '+
          '<a href="#" data-bretry style="color:#E51652;font-weight:700">Try again</a> '+
          'or call us on 0161 833 5400.';
        var a=panel.querySelector('[data-bretry]');
        if(a) a.addEventListener('click',function(ev){
          ev.preventDefault();
          msgEl.textContent=msgs[0]; mi=0;
          rot=setInterval(function(){ mi=(mi+1)%msgs.length; msgEl.textContent=msgs[mi]; },5000);
          var src=f.src; f.src='about:blank';
          setTimeout(function(){ f.src=src; },60);
          arm();
        });
      },25000);
    };
    arm();
  }
  function appReady(){
    if(ready) return; ready=true;
    if(rot) clearInterval(rot);
    if(tm) clearTimeout(tm);
    if(panel&&panel.parentNode) panel.parentNode.removeChild(panel);
  }
  s.parentNode.insertBefore(wrap, s);
  window.addEventListener('message', function(e){
    if(!e.data) return;
    // The first height report is the app saying "I have rendered" — replace
    // the loading panel at once, never prolong it (spec: never delay merely
    // to display all messages).
    if(e.data.brauditHeight){ f.style.height = e.data.brauditHeight+'px'; appReady(); }
    if(e.data.brauditReady) appReady();
    // Screen changes and the register-search overlay live INSIDE a very tall
    // iframe; window.scrollTo in there is a no-op and position:fixed pins to
    // the iframe box, not the visitor's viewport. Without this the search
    // wait looked like a crash on WordPress (Jonathan, 26 Aug): the hint
    // carousel was rendering 1000px+ below the fold. The app asks, we scroll.
    if(e.data.brauditScrollTop){
      try{ wrap.scrollIntoView({behavior:'smooth', block:'start'}); }catch(_e){}
    }
    // The compact box hands the visitor over to the full wizard. It cannot
    // navigate the host page itself from inside a cross-origin iframe, so it
    // asks us to. Only ever a navigation, never arbitrary script.
    if(e.data.brauditNavigate) window.location.href = e.data.brauditNavigate;
  });
})();"""


# ── static assets ────────────────────────────────────────────────────────
# Added 17 Aug 2026. Until now the engine served only the pages in _PAGES and
# embed.js — there was no route for a stylesheet or an image, which is why the
# six widgets each carried a private copy of the CSS and why brand/*.png could
# not load anywhere.
#
# Only these three directories, and only these extensions. This is a public,
# unauthenticated host, so the path is treated as hostile:
#
#   * the resolved real path must sit INSIDE the served directory, checked
#     with os.path.commonpath after realpath — that defeats ../ traversal,
#     encoded traversal, and a symlink pointing out of the tree, which a
#     simple '..' not in path check does not
#   * unknown extensions are refused rather than served as octet-stream, so a
#     stray .py or .env in web/ can never be downloaded
_ASSET_DIRS = {'/braudit.css': ('', 'braudit.css'),
               '/brand/': ('brand', None),
               '/fonts/': ('fonts', None)}
_ASSET_TYPES = {
    '.css': 'text/css; charset=utf-8',
    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
    '.gif': 'image/gif', '.svg': 'image/svg+xml', '.webp': 'image/webp',
    '.ico': 'image/x-icon',
    '.woff2': 'font/woff2', '.woff': 'font/woff',
}


def _static(path: str):
    """Return (bytes|None, content_type) for a static asset path, else None.

    None  -> not a static route at all, let the caller carry on routing.
    (None, ct) -> it IS a static route but the file isn't there: a real 404.
    """
    if path == '/braudit.css':
        rel = 'braudit.css'
    elif path.startswith('/brand/') or path.startswith('/fonts/'):
        rel = path.lstrip('/')
    else:
        return None

    ext = os.path.splitext(rel)[1].lower()
    ctype = _ASSET_TYPES.get(ext)
    if not ctype:
        return (None, 'text/plain')          # unknown type -> 404, never served

    full = os.path.realpath(os.path.join(_WEB, rel))
    root = os.path.realpath(_WEB)
    # commonpath raises ValueError on different drives/roots; treat as outside.
    try:
        inside = os.path.commonpath([full, root]) == root
    except ValueError:
        inside = False
    if not inside or not os.path.isfile(full):
        return (None, ctype)
    try:
        with open(full, 'rb') as f:
            return (f.read(), ctype)
    except OSError:
        return (None, ctype)


def _audit_pay(payload: dict) -> dict:
    """Create a Stripe Checkout session for the audit checkout's
    "over the phone now" option (Jonathan, 27 Aug).

    The AMOUNT IS COMPUTED HERE, never trusted from the browser: £99 net
    always (the Audit Promotion discount is applied upstream in the page's
    display), plus 20% VAT unless the applicant is a non-UK organisation or
    non-UK individual (vat_exempt). The consultation is free either way, so
    consult only changes the line-item description, not the price.

    Needs STRIPE_SECRET_KEY in the environment (Render env var — the value
    lives in temmy-access/secrets.env, never in this repo). Unset -> the
    page falls back to "we'll call you to take payment".
    """
    import urllib.error
    import urllib.parse
    import urllib.request

    sk = (os.environ.get('STRIPE_SECRET_KEY') or '').strip()
    if not sk:
        return {'ok': False, 'error': 'payment not configured', 'status': 200}

    vat_exempt = bool(payload.get('vat_exempt'))
    consult = bool(payload.get('consult'))
    ref = str(payload.get('ref') or '')[:40]
    email = str(payload.get('email') or '')[:200]
    net_p = 9900                                   # £99.00 net, always
    total = net_p if vat_exempt else int(round(net_p * 1.20))
    name = ('Trademark Audit & Consultation' if consult else 'Trademark Audit')
    desc = 'Audit Promotion applied' + (' — VAT not applicable (outside UK)'
                                        if vat_exempt else ' — includes VAT @ 20%')
    form = {
        'mode': 'payment',
        'client_reference_id': ref or 'AUD',
        'success_url': 'https://www.thetrademarkhelpline.com/?audit_paid=1',
        'cancel_url': 'https://www.thetrademarkhelpline.com/?audit_paid=0',
        'line_items[0][quantity]': '1',
        'line_items[0][price_data][currency]': 'gbp',
        'line_items[0][price_data][unit_amount]': str(total),
        'line_items[0][price_data][product_data][name]': name,
        'line_items[0][price_data][product_data][description]': desc,
        'metadata[invoice_ref]': ref,
        'metadata[session_id]': str(payload.get('session_id') or '')[:60],
    }
    if email:
        form['customer_email'] = email
    req = urllib.request.Request(
        'https://api.stripe.com/v1/checkout/sessions',
        data=urllib.parse.urlencode(form).encode(),
        headers={'Authorization': 'Bearer ' + sk,
                 'Content-Type': 'application/x-www-form-urlencoded'})
    try:
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read().decode())
        return {'ok': True, 'url': data.get('url'), 'status': 200}
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read().decode()).get('error', {}).get('message', '')
        except Exception:
            err = str(e)
        return {'ok': False, 'error': f'stripe: {err[:200]}', 'status': 200}
    except Exception as e:
        return {'ok': False, 'error': f'stripe unreachable: {e}', 'status': 200}


def _allowed_origin(origin: str) -> str:
    """CORS allow-list from env. ALLOWED_ORIGINS unset -> '*' (dev only);
    set it to a comma-separated list in production (TMH, portal, partners)."""
    allowed = (os.environ.get('ALLOWED_ORIGINS') or '').strip()
    if not allowed:
        return '*'
    if origin and origin in [o.strip() for o in allowed.split(',')]:
        return origin
    return 'null'


class _Handler(BaseHTTPRequestHandler):
    server_version = 'BrauditFreeSearch/1.0'

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin',
                         _allowed_origin(self.headers.get('Origin', '')))
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, X-Engine-Key')

    def _send(self, body: dict, status: int = 200):
        payload = json.dumps(body).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self._cors()
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _send_raw(self, payload: bytes, ctype: str, status: int = 200):
        self.send_response(status)
        self.send_header('Content-Type', ctype)
        self._cors()
        # Allow partner pages to iframe the widget.
        self.send_header('Content-Security-Policy', 'frame-ancestors *')
        # Added 10 Aug 2026. These responses previously carried NO cache
        # headers, so browsers fell back to heuristic caching and could keep
        # serving an old copy of the widget for hours after a deploy — which
        # is how a fixed bug appears to still be broken, and how a partner
        # site can sit on stale markup indefinitely.
        #
        # no-cache does NOT mean "never store": it means "always revalidate
        # before use". The browser still keeps the file and we still get a
        # cheap 304 when nothing changed. These are small single files served
        # from one host, so the cost is a round trip, and the alternative is
        # visitors running whichever version they happened to load first.
        self.send_header('Cache-Control', 'no-cache, must-revalidate')
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _engine_key_ok(self) -> bool:
        """True unless a key is configured and the caller did not present it."""
        want = (os.environ.get('ENGINE_SHARED_KEY') or '').strip()
        if not want:
            return True                       # not configured — open, as before
        got = (self.headers.get('X-Engine-Key') or '').strip()
        # Constant-time compare: a plain == leaks the key one character at a
        # time to anyone patient enough to measure the response.
        import hmac
        return hmac.compare_digest(got, want)

    def do_OPTIONS(self):
        self._send({}, 204)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip('/')
        params = {k: v[0] for k, v in parse_qs(parsed.query).items()}
        if path in _PAGES or path.startswith('/report/'):
            # /report/<session-uuid> is the pretty unique-URL form of the
            # search report; the page reads the id from the path itself.
            page = os.path.join(_WEB, _PAGES.get(path, 'report.html'))
            try:
                with open(page, 'rb') as f:
                    self._send_raw(f.read(), 'text/html; charset=utf-8')
            except OSError:
                self._send({'ok': False, 'error': 'widget not found'}, 404)
            return
        if path == '/embed.js':
            self._send_raw(_EMBED_JS.encode(), 'application/javascript')
            return
        asset = _static(parsed.path)
        if asset is not None:
            body, ctype = asset
            if body is None:
                self._send({'ok': False, 'error': 'not found'}, 404)
            else:
                self._send_raw(body, ctype)
            return
        if path == '/healthz':
            self._send({'ok': True})
            return
        if path == '/nuggets':
            from .nuggets import payload as _nug
            self._send({'ok': True, 'nuggets': _nug()})
            return
        if path == '/jurisdictions':
            out = handle_jurisdictions()
            self._send(out, out.get('status', 200))
        elif path.startswith('/lookup/'):
            action = path.split('/lookup/', 1)[1]
            out = handle_lookup(action, params, _make_client())
            self._send(out, out.get('status', 200))
        elif path in ('', '/health'):
            self._send({'ok': True, 'service': 'free-search'})
        else:
            self._send({'ok': False, 'error': 'not found'}, 404)

    def do_POST(self):
        path = self.path.rstrip('/')
        if path not in ('/free-search', '/enrich', '/suggest-classes', '/read-website',
                        '/class-scope', '/audit-pay'):
            self._send({'ok': False, 'error': 'not found'}, 404)
            return
        if not self._engine_key_ok():
            # The usage gate lives in the Supabase Edge Function, in FRONT of
            # this engine. Without a shared secret that gate is decorative:
            # anyone who opens developer tools sees this URL and can call it
            # directly, for ever, for free — and the AI routes cost real money
            # per call.
            #
            # Opt-in by design. Unset ENGINE_SHARED_KEY and nothing changes,
            # so deploying this cannot break a live site; set it here and on
            # the Edge Function and the back door closes. Set it AFTER both
            # are deployed, never between.
            self._send({'ok': False, 'error': 'forbidden'}, 403)
            return
        try:
            length = int(self.headers.get('Content-Length', 0))
            payload = json.loads(self.rfile.read(length) or b'{}')
        except (ValueError, json.JSONDecodeError):
            self._send({'ok': False, 'error': 'invalid JSON'}, 400)
            return
        if path == '/audit-pay':
            out = _audit_pay(payload)
        elif path == '/enrich':
            out = handle_enrich(payload)
        elif path == '/suggest-classes':
            out = handle_suggest_classes(payload)
        elif path == '/read-website':
            out = handle_read_website(payload)
        elif path == '/class-scope':
            # Class Builder's review screen: the application-grade selection
            # (official Nice headings + registered-vocabulary terms) WITHOUT
            # running a register search. Same spec_terms source of truth as
            # the search journey's scope capture.
            try:
                try:
                    from . import spec_terms  # package context (Render)
                except ImportError:
                    import spec_terms         # bare-script context (local dev)
                rows = spec_terms.build_application_scope(
                    payload.get('classes'), payload.get('class_source'))
                out = {'ok': True, 'application_scope': rows}
            except Exception as e:  # degrade loudly but JSON-shaped
                out = {'ok': False, 'error': f'class-scope failed: {e}',
                       'status': 500}
        else:
            out = handle_free_search(payload, _make_client())
        self._send(out, out.get('status', 200))

    def log_message(self, *args):  # keep dev-server output quiet
        pass


def serve(host: str = '0.0.0.0', port: int = 8080):
    srv = ThreadingHTTPServer((host, port), _Handler)
    print(f'Free Search dev server on http://{host}:{port}')
    print('  GET  /jurisdictions   POST /free-search   POST /enrich')
    srv.serve_forever()


if __name__ == '__main__':
    serve(port=int(os.environ.get('PORT', 8080)))
