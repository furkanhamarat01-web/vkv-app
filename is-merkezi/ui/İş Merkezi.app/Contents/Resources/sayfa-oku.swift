import AppKit
import Foundation
import WebKit

final class Reader: NSObject, WKNavigationDelegate {
    let view: WKWebView
    let url: URL
    var finished = false
    var expanded = false
    var attempts = 0

    init(url: URL) {
        self.url = url
        let config = WKWebViewConfiguration()
        config.websiteDataStore = .nonPersistent()
        view = WKWebView(frame: NSRect(x: 0, y: 0, width: 1280, height: 900), configuration: config)
        super.init()
        view.customUserAgent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15"
        view.navigationDelegate = self
    }

    func start() {
        var request = URLRequest(url: url)
        request.setValue("tr-TR,tr;q=0.9,en;q=0.6", forHTTPHeaderField: "Accept-Language")
        view.load(request)
        DispatchQueue.main.asyncAfter(deadline: .now() + 30) { self.collect(final: true) }
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        DispatchQueue.main.asyncAfter(deadline: .now() + 2.5) { self.expand() }
    }

    func webView(_ webView: WKWebView, didFail navigation: WKNavigation!, withError error: Error) {
        emit(["error": error.localizedDescription])
    }

    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        emit(["error": error.localizedDescription])
    }

    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction, decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        decisionHandler(expanded && action.targetFrame?.isMainFrame == true && action.navigationType == .linkActivated ? .cancel : .allow)
    }

    func expand() {
        if finished || expanded { return }
        expanded = true
        let script = """
        (() => {
          const clean = s => (s || "").replace(/\\s+/g, " ").trim();
          const label = /^(teknik özellikler|ürün özellikleri|tüm özellikler|özellikler|teknik detaylar|specifications|technical specifications|tech specs)$/i;
          let clicked = 0;
          document.querySelectorAll("button,[role=tab],[role=button],h2,h3,li,span,div,a").forEach(el => {
            if (clicked >= 4 || !label.test(clean(el.textContent))) return;
            const link = el.closest("a[href]");
            if (link && !link.getAttribute("href").startsWith("#")) return;
            try { el.click(); clicked += 1; } catch (_) {}
          });
          return clicked;
        })()
        """
        view.evaluateJavaScript(script) { _, _ in
            DispatchQueue.main.asyncAfter(deadline: .now() + 2) { self.collect(final: false) }
        }
    }

    func collect(final: Bool) {
        if finished { return }
        let script = """
        (() => {
          const ld = [...document.querySelectorAll('script[type="application/ld+json"]')].map(s => s.textContent).slice(0, 12);
          const meta = n => (document.querySelector(`meta[property="${n}"],meta[name="${n}"],meta[itemprop="${n}"]`) || {}).content || "";
          const clean = s => (s || "").replace(/\\s+/g, " ").trim();
          const specs = [];
          const seen = new Set();
          const add = (k, v) => {
            k = clean(k).replace(/:$/, "");
            v = clean(v);
            if (!k || !v || k.length > 60 || v.length > 200 || k === v || specs.length >= 150) return;
            if (/^\\d+([.,]\\d+)?\\s*\\(\\d+\\)$/.test(v) || /^[\\d.,]+$/.test(k)) return;
            const key = k.toLocaleLowerCase("tr");
            if (seen.has(key)) return;
            seen.add(key);
            specs.push([k, v]);
          };
          const unescape = s => { try { return JSON.parse('"' + s + '"'); } catch (_) { return s; } };
          const raw = document.documentElement.innerHTML;
          const pair = /"(?:name|key|label|attributeName)"\\s*:\\s*"((?:[^"\\\\]|\\\\.){2,60})"\\s*,\\s*"(?:property|value|unitText|attributeValue|values?)"\\s*:\\s*"((?:[^"\\\\]|\\\\.){1,200})"/g;
          const jsonSpecs = [];
          let m;
          while ((m = pair.exec(raw)) && jsonSpecs.length < 300) jsonSpecs.push([unescape(m[1]), unescape(m[2])]);
          const skip = el => el.closest("nav,footer,header,[class*=recommend],[class*=reklam],[class*=carousel],[class*=slider],[class*=review],[class*=rating],[class*=comment]");
          const leaf = el => !el.querySelector("div,li,tr,dl,ul,table");
          document.querySelectorAll("tr").forEach(tr => {
            if (skip(tr)) return;
            const cells = [...tr.children].filter(e => e.tagName === "TD" || e.tagName === "TH");
            if (cells.length === 2) add(cells[0].textContent, cells[1].textContent);
          });
          document.querySelectorAll("dt").forEach(dt => {
            const dd = dt.nextElementSibling;
            if (dd && dd.tagName === "DD" && !skip(dt)) add(dt.textContent, dd.textContent);
          });
          document.querySelectorAll("div,ul,section").forEach(group => {
            const rows = [...group.children];
            if (rows.length < 4 || skip(group)) return;
            const pairs = rows.map(row => [...row.children].filter(k => clean(k.textContent))).filter(kids => kids.length === 2 && leaf(kids[0]) && clean(kids[0].textContent).length <= 60);
            if (pairs.length < 4 || pairs.length < rows.length * 0.6) return;
            const numeric = pairs.filter(kids => /^[\\d.,\\s()]+$/.test(clean(kids[1].textContent))).length;
            if (numeric >= pairs.length * 0.8) return;
            pairs.forEach(kids => add(kids[0].textContent, kids[1].textContent));
          });
          return JSON.stringify({
            specs: specs,
            jsonSpecs: jsonSpecs,
            title: document.title || "",
            ogTitle: meta("og:title"),
            description: meta("og:description") || meta("description"),
            price: meta("product:price:amount") || meta("og:price:amount") || meta("price"),
            ld: ld,
            text: (document.body ? document.body.innerText : "").slice(0, 20000),
            url: location.href
          });
        })()
        """
        view.evaluateJavaScript(script) { value, error in
            if self.finished { return }
            guard let raw = value as? String,
                  let data = raw.data(using: .utf8),
                  let object = try? JSONSerialization.jsonObject(with: data) as? [String: Any] else {
                if final { self.emit(["error": error?.localizedDescription ?? "okunamadı"]) }
                return
            }
            let text = (object["text"] as? String) ?? ""
            let letters = text.filter { $0.isLetter }.count
            let specs = max((object["specs"] as? [Any])?.count ?? 0, (object["jsonSpecs"] as? [Any])?.count ?? 0)
            if !final && (letters < 400 || specs < 4) && self.attempts < 5 {
                self.attempts += 1
                DispatchQueue.main.asyncAfter(deadline: .now() + 2) { self.collect(final: false) }
                return
            }
            self.emit(object)
        }
    }

    func emit(_ object: [String: Any]) {
        if finished { return }
        finished = true
        if let data = try? JSONSerialization.data(withJSONObject: object),
           let text = String(data: data, encoding: .utf8) {
            FileHandle.standardOutput.write(text.data(using: .utf8)!)
        }
        exit(0)
    }
}

guard CommandLine.arguments.count == 2,
      let url = URL(string: CommandLine.arguments[1]),
      let scheme = url.scheme?.lowercased(), scheme == "http" || scheme == "https" else {
    FileHandle.standardError.write("adres gerekli\n".data(using: .utf8)!)
    exit(2)
}

let app = NSApplication.shared
app.setActivationPolicy(.prohibited)
let reader = Reader(url: url)
reader.start()
app.run()
