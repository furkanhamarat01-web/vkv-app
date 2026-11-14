import AppKit
import Foundation
import PDFKit

struct Job: Decodable {
    let output: String
    let party: String
    let meta: String
    let reason: String
    let files: [String]
}

let pageSize = CGSize(width: 595, height: 842)
let job = try JSONDecoder().decode(Job.self, from: Data(contentsOf: URL(fileURLWithPath: CommandLine.arguments[1])))
let document = PDFDocument()

func addImage(_ image: NSImage) {
    guard let page = PDFPage(image: image) else { return }
    document.insert(page, at: document.pageCount)
}

func render(_ draw: @escaping (CGRect) -> Void) -> NSImage {
    NSImage(size: pageSize, flipped: true) { rect in
        NSColor.white.setFill()
        rect.fill()
        draw(rect)
        return true
    }
}

func addWrapped(title: String, body: String) {
    let width = pageSize.width - 96
    let titleFont = NSFont.systemFont(ofSize: 18, weight: .bold)
    let bodyFont = NSFont.systemFont(ofSize: 12)
    let titleHeight = title.isEmpty ? 0 : (title as NSString).boundingRect(
        with: CGSize(width: width, height: 200),
        options: [.usesLineFragmentOrigin],
        attributes: [.font: titleFont]
    ).height + 16
    let storage = NSTextStorage(string: body, attributes: [
        .font: bodyFont,
        .foregroundColor: NSColor.black,
    ])
    let manager = NSLayoutManager()
    storage.addLayoutManager(manager)
    var made = false
    while manager.textContainers.isEmpty || manager.glyphRange(for: manager.textContainers.last!).length > 0 && NSMaxRange(manager.glyphRange(for: manager.textContainers.last!)) < manager.numberOfGlyphs {
        let height = pageSize.height - 96 - (manager.textContainers.isEmpty ? titleHeight : 0)
        let container = NSTextContainer(size: CGSize(width: width, height: max(120, height)))
        container.lineFragmentPadding = 0
        manager.addTextContainer(container)
        let range = manager.glyphRange(for: container)
        if range.length == 0 { break }
        let showTitle = !made && !title.isEmpty
        let image = render { _ in
            if showTitle {
                (title as NSString).draw(
                    in: NSRect(x: 48, y: 40, width: width, height: titleHeight),
                    withAttributes: [.font: titleFont, .foregroundColor: NSColor.black]
                )
            }
            let originY: CGFloat = showTitle ? 40 + titleHeight : 48
            manager.drawGlyphs(forGlyphRange: range, at: NSPoint(x: 48, y: originY))
        }
        addImage(image)
        made = true
        if NSMaxRange(range) >= manager.numberOfGlyphs { break }
    }
    if !made {
        addImage(render { _ in
            (title as NSString).draw(
                in: NSRect(x: 48, y: 48, width: width, height: 80),
                withAttributes: [.font: titleFont, .foregroundColor: NSColor.black]
            )
        })
    }
}

func addFitted(_ source: NSImage, caption: String) {
    let image = render { rect in
        (caption as NSString).draw(
            at: NSPoint(x: 36, y: 24),
            withAttributes: [.font: NSFont.systemFont(ofSize: 11, weight: .semibold), .foregroundColor: NSColor.darkGray]
        )
        let box = rect.insetBy(dx: 36, dy: 36).offsetBy(dx: 0, dy: 16)
        let size = source.size
        guard size.width > 0, size.height > 0 else { return }
        let scale = min(box.width / size.width, box.height / size.height)
        let fitted = NSSize(width: size.width * scale, height: size.height * scale)
        source.draw(in: NSRect(
            x: box.midX - fitted.width / 2,
            y: box.midY - fitted.height / 2,
            width: fitted.width,
            height: fitted.height
        ))
    }
    addImage(image)
}

struct Mime {
    var headers: [String: String]
    var body: Data
    var children: [Mime]
}

struct MailFile {
    var name: String
    var data: Data
    var type: String
}

struct MailPiece {
    var text: String
    var files: [MailFile]
}

func replaceText(_ text: String, _ pattern: String, _ template: String) -> String {
    guard let expression = try? NSRegularExpression(pattern: pattern) else { return text }
    let range = NSRange(location: 0, length: (text as NSString).length)
    return expression.stringByReplacingMatches(in: text, range: range, withTemplate: template)
}

func headerParameter(_ header: String, _ name: String) -> String? {
    let pattern = "(?i)(?:^|;)\\s*\(NSRegularExpression.escapedPattern(for: name))\\s*=\\s*(?:\"([^\"]*)\"|([^;\\s]+))"
    guard let expression = try? NSRegularExpression(pattern: pattern) else { return nil }
    let ns = header as NSString
    guard let match = expression.firstMatch(in: header, range: NSRange(location: 0, length: ns.length)) else { return nil }
    if match.range(at: 1).location != NSNotFound { return ns.substring(with: match.range(at: 1)) }
    if match.range(at: 2).location != NSNotFound { return ns.substring(with: match.range(at: 2)) }
    return nil
}

func charsetEncoding(_ name: String) -> String.Encoding {
    switch name.lowercased().replacingOccurrences(of: "_", with: "-") {
    case "utf-8", "utf8":
        return .utf8
    case "iso-8859-9", "latin5", "windows-1254", "cp1254":
        return .windowsCP1254
    case "windows-1252", "cp1252":
        return .windowsCP1252
    case "iso-8859-1", "latin1", "us-ascii", "ascii":
        return .isoLatin1
    default:
        return .utf8
    }
}

func hexNibble(_ byte: UInt8) -> UInt8? {
    switch byte {
    case 48...57: return byte - 48
    case 65...70: return byte - 55
    case 97...102: return byte - 87
    default: return nil
    }
}

func decodeQuotedPrintable(_ data: Data) -> Data {
    let bytes = [UInt8](data)
    var out = Data()
    var index = 0
    while index < bytes.count {
        if bytes[index] == 61 {
            let next = index + 1
            if next < bytes.count && (bytes[next] == 10 || bytes[next] == 13) {
                index = next + 1
                if index < bytes.count && bytes[index - 1] == 13 && bytes[index] == 10 { index += 1 }
                continue
            }
            if index + 2 < bytes.count, let high = hexNibble(bytes[index + 1]), let low = hexNibble(bytes[index + 2]) {
                out.append((high << 4) | low)
                index += 3
                continue
            }
        }
        if bytes[index] != 13 { out.append(bytes[index]) }
        index += 1
    }
    return out
}

func decodeEncodedWord(charset: String, kind: String, payload: String) -> String {
    let encoding = charsetEncoding(charset)
    if kind == "b" {
        let data = Data(base64Encoded: payload, options: .ignoreUnknownCharacters) ?? Data()
        return String(data: data, encoding: encoding) ?? String(data: data, encoding: .utf8) ?? payload
    }
    var bytes = Data()
    let chars = Array(payload.utf8)
    var index = 0
    while index < chars.count {
        if chars[index] == 95 {
            bytes.append(32)
            index += 1
        } else if chars[index] == 61, index + 2 < chars.count, let high = hexNibble(chars[index + 1]), let low = hexNibble(chars[index + 2]) {
            bytes.append((high << 4) | low)
            index += 3
        } else {
            bytes.append(chars[index])
            index += 1
        }
    }
    return String(data: bytes, encoding: encoding) ?? String(data: bytes, encoding: .utf8) ?? payload
}

func decodeHeaderWords(_ value: String) -> String {
    let pattern = "=\\?([^?]+)\\?([bBqQ])\\?([^?]*)\\?="
    guard let expression = try? NSRegularExpression(pattern: pattern) else { return value }
    let ns = value as NSString
    let matches = expression.matches(in: value, range: NSRange(location: 0, length: ns.length))
    if matches.isEmpty { return value }
    var out = ""
    var cursor = 0
    var lastWasWord = false
    for match in matches {
        let gap = ns.substring(with: NSRange(location: cursor, length: match.range.location - cursor))
        if !(lastWasWord && gap.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty) {
            out += gap
        }
        out += decodeEncodedWord(
            charset: ns.substring(with: match.range(at: 1)),
            kind: ns.substring(with: match.range(at: 2)).lowercased(),
            payload: ns.substring(with: match.range(at: 3))
        )
        cursor = match.range.location + match.range.length
        lastWasWord = true
    }
    out += ns.substring(from: cursor)
    return out
}

func headerMap(_ data: Data) -> [String: String] {
    let raw = (String(data: data, encoding: .isoLatin1) ?? "")
        .replacingOccurrences(of: "\r\n", with: "\n")
        .replacingOccurrences(of: "\r", with: "\n")
    var lines: [String] = []
    for line in raw.split(separator: "\n", omittingEmptySubsequences: false) {
        let text = String(line)
        if (text.hasPrefix(" ") || text.hasPrefix("\t")), !lines.isEmpty {
            lines[lines.count - 1] += " " + text.trimmingCharacters(in: .whitespaces)
        } else {
            lines.append(text)
        }
    }
    var map: [String: String] = [:]
    for line in lines {
        guard let colon = line.firstIndex(of: ":") else { continue }
        let key = line[..<colon].trimmingCharacters(in: .whitespaces).lowercased()
        if key.isEmpty || map[key] != nil { continue }
        let value = line[line.index(after: colon)...].trimmingCharacters(in: .whitespaces)
        map[key] = decodeHeaderWords(String(value))
    }
    return map
}

func splitHeaderBody(_ data: Data) -> (Data, Data)? {
    let crlf = Data([13, 10, 13, 10])
    let lf = Data([10, 10])
    if let range = data.range(of: crlf) {
        return (data.subdata(in: 0..<range.lowerBound), data.subdata(in: range.upperBound..<data.endIndex))
    }
    if let range = data.range(of: lf) {
        return (data.subdata(in: 0..<range.lowerBound), data.subdata(in: range.upperBound..<data.endIndex))
    }
    return nil
}

func mediaType(_ header: String?) -> String {
    let raw = header ?? "text/plain"
    let token = raw.split(separator: ";", maxSplits: 1).first.map(String.init) ?? "text/plain"
    let type = token.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
    return type.isEmpty ? "text/plain" : type
}

func splitParts(_ body: Data, boundary: String) -> [Data] {
    guard let marker = ("--" + boundary).data(using: .utf8) else { return [] }
    var parts: [Data] = []
    var search = body.startIndex
    var start: Data.Index?
    while let range = body.range(of: marker, in: search..<body.endIndex) {
        if let partStart = start {
            var chunkEnd = range.lowerBound
            if chunkEnd >= body.startIndex + 2, body[chunkEnd - 2] == 13, body[chunkEnd - 1] == 10 {
                chunkEnd -= 2
            } else if chunkEnd > body.startIndex, body[chunkEnd - 1] == 10 {
                chunkEnd -= 1
            }
            if chunkEnd >= partStart { parts.append(body.subdata(in: partStart..<chunkEnd)) }
        }
        var index = range.upperBound
        if index + 1 < body.endIndex, body[index] == 45, body[index + 1] == 45 { break }
        while index < body.endIndex, body[index] == 32 || body[index] == 9 { index += 1 }
        if index < body.endIndex, body[index] == 13 { index += 1 }
        if index < body.endIndex, body[index] == 10 { index += 1 }
        start = index
        search = index
    }
    return parts
}

func parseMime(_ data: Data, depth: Int = 0) -> Mime? {
    guard depth < 6, let (headerData, body) = splitHeaderBody(data) else { return nil }
    let headers = headerMap(headerData)
    let type = mediaType(headers["content-type"])
    if type.hasPrefix("multipart/"), let boundary = headerParameter(headers["content-type"] ?? "", "boundary"), !boundary.isEmpty {
        let children = splitParts(body, boundary: boundary).compactMap { parseMime($0, depth: depth + 1) }
        return Mime(headers: headers, body: Data(), children: children)
    }
    if type == "message/rfc822", let child = parseMime(body, depth: depth + 1) {
        return Mime(headers: headers, body: Data(), children: [child])
    }
    return Mime(headers: headers, body: body, children: [])
}

func transferDecoded(_ mime: Mime) -> Data {
    let encoding = (mime.headers["content-transfer-encoding"] ?? "").lowercased()
    if encoding.contains("quoted-printable") { return decodeQuotedPrintable(mime.body) }
    if encoding.contains("base64") {
        let text = String(data: mime.body, encoding: .isoLatin1) ?? ""
        return Data(base64Encoded: text, options: .ignoreUnknownCharacters) ?? Data()
    }
    return Data(mime.body.filter { $0 != 13 })
}

func decodedText(_ mime: Mime) -> String {
    let data = transferDecoded(mime)
    let charset = headerParameter(mime.headers["content-type"] ?? "", "charset") ?? "utf-8"
    return String(data: data, encoding: charsetEncoding(charset))
        ?? String(data: data, encoding: .utf8)
        ?? String(data: data, encoding: .isoLatin1)
        ?? ""
}

func partName(_ mime: Mime) -> String {
    let disposition = mime.headers["content-disposition"] ?? ""
    let type = mime.headers["content-type"] ?? ""
    if let extended = headerParameter(disposition, "filename*") ?? headerParameter(type, "name*") {
        let pieces = extended.split(separator: "'", maxSplits: 2, omittingEmptySubsequences: false)
        let encoded = pieces.count == 3 ? String(pieces[2]) : extended
        return encoded.removingPercentEncoding ?? encoded
    }
    return headerParameter(disposition, "filename") ?? headerParameter(type, "name") ?? ""
}

func decodeEntities(_ text: String) -> String {
    let named = [
        "amp": "&", "lt": "<", "gt": ">", "quot": "\"", "apos": "'", "nbsp": " ", "shy": "",
        "ccedil": "ç", "ouml": "ö", "uuml": "ü", "auml": "ä", "euml": "ë", "iuml": "ï",
        "scedil": "ş", "gbreve": "ğ", "inodot": "ı", "szlig": "ß", "ntilde": "ñ",
        "agrave": "à", "aacute": "á", "acirc": "â", "atilde": "ã", "aring": "å", "aelig": "æ",
        "egrave": "è", "eacute": "é", "ecirc": "ê", "igrave": "ì", "iacute": "í", "icirc": "î",
        "ograve": "ò", "oacute": "ó", "ocirc": "ô", "otilde": "õ", "ugrave": "ù", "uacute": "ú",
        "ucirc": "û", "yacute": "ý", "yuml": "ÿ", "eth": "ð", "thorn": "þ",
        "ndash": "–", "mdash": "—", "lsquo": "‘", "rsquo": "’", "ldquo": "“", "rdquo": "”",
        "bull": "•", "hellip": "…", "euro": "€", "trade": "™", "deg": "°", "middot": "·",
        "laquo": "«", "raquo": "»", "copy": "©", "reg": "®",
    ]
    let pattern = "&(#x[0-9a-fA-F]+|#\\d+|[a-zA-Z][a-zA-Z0-9]+);"
    guard let expression = try? NSRegularExpression(pattern: pattern) else { return text }
    let ns = text as NSString
    var out = ""
    var cursor = 0
    for match in expression.matches(in: text, range: NSRange(location: 0, length: ns.length)) {
        out += ns.substring(with: NSRange(location: cursor, length: match.range.location - cursor))
        let token = ns.substring(with: match.range(at: 1))
        if token.hasPrefix("#x"), let value = Int(token.dropFirst(2), radix: 16), let scalar = UnicodeScalar(value) {
            out.append(Character(scalar))
        } else if token.hasPrefix("#"), let value = Int(token.dropFirst()), let scalar = UnicodeScalar(value) {
            out.append(Character(scalar))
        } else {
            out += named[token.lowercased()] ?? ns.substring(with: match.range)
        }
        cursor = match.range.location + match.range.length
    }
    out += ns.substring(from: cursor)
    return out
}

func tidyText(_ text: String) -> String {
    let normalized = text.replacingOccurrences(of: "\r\n", with: "\n").replacingOccurrences(of: "\r", with: "\n")
    var lines: [String] = []
    var blank = false
    for raw in normalized.components(separatedBy: "\n") {
        let cells = raw.components(separatedBy: "\t").map { cell in
            cell.replacingOccurrences(of: "\u{00A0}", with: " ")
                .split(whereSeparator: { $0.isWhitespace })
                .joined(separator: " ")
        }.filter { !$0.isEmpty }
        let line = cells.joined(separator: "  |  ")
        if line.isEmpty {
            if !blank && !lines.isEmpty { lines.append("") }
            blank = true
        } else {
            lines.append(line)
            blank = false
        }
    }
    while lines.last == "" { lines.removeLast() }
    return lines.joined(separator: "\n")
}

func htmlToText(_ html: String) -> String {
    var text = html.replacingOccurrences(of: "\r\n", with: "\n").replacingOccurrences(of: "\r", with: "\n")
    text = replaceText(text, "(?is)<(script|style|head)[^>]*>.*?</\\1>", " ")
    text = replaceText(text, "(?s)<!--.*?-->", " ")
    text = replaceText(text, ">\\s+<", "><")
    text = replaceText(text, "(?i)<br\\s*/?>", "\n")
    text = replaceText(text, "(?i)</(p|div|tr|h[1-6]|li|table|section|blockquote|ul|ol)>", "\n")
    text = replaceText(text, "(?i)<(p|h[1-6]|li|blockquote)(?:\\s[^>]*)?>", "\n")
    text = replaceText(text, "(?i)</t[dh]>", "\t")
    text = replaceText(text, "<[^>]+>", "")
    return tidyText(decodeEntities(text))
}

func gatherMail(_ mime: Mime) -> MailPiece {
    let type = mediaType(mime.headers["content-type"])
    if type == "multipart/alternative" {
        var best = MailPiece(text: "", files: [])
        for child in mime.children {
            let piece = gatherMail(child)
            if piece.text.count > best.text.count { best = piece }
        }
        if best.files.isEmpty {
            best.files = mime.children.flatMap { gatherMail($0).files }
        }
        return best
    }
    if type.hasPrefix("multipart/") || type == "message/rfc822" {
        var text: [String] = []
        var files: [MailFile] = []
        for child in mime.children {
            let piece = gatherMail(child)
            if !piece.text.isEmpty { text.append(piece.text) }
            files.append(contentsOf: piece.files)
        }
        return MailPiece(text: text.joined(separator: "\n\n"), files: files)
    }
    let disposition = (mime.headers["content-disposition"] ?? "").lowercased()
    let name = partName(mime)
    if type == "text/plain" || type == "text/html" {
        if disposition.contains("attachment") {
            return MailPiece(text: name.isEmpty ? "" : "Ek: \(name)", files: [])
        }
        let raw = type == "text/html" ? htmlToText(decodedText(mime)) : tidyText(decodedText(mime))
        return MailPiece(text: raw, files: [])
    }
    if type == "application/pdf" || type.hasPrefix("image/") {
        let data = transferDecoded(mime)
        guard !data.isEmpty else { return MailPiece(text: "", files: []) }
        return MailPiece(text: "", files: [MailFile(name: name.isEmpty ? "ek" : name, data: data, type: type)])
    }
    if disposition.contains("attachment"), !name.isEmpty {
        return MailPiece(text: "Ek: \(name)", files: [])
    }
    return MailPiece(text: "", files: [])
}

func oneLine(_ value: String, limit: Int) -> String {
    let line = value.split(whereSeparator: { $0.isWhitespace }).joined(separator: " ")
    if line.count <= limit { return line }
    return String(line.prefix(limit))
}

func addMailFile(_ file: MailFile) {
    if file.type == "application/pdf" || file.data.starts(with: Data("%PDF-".utf8)) {
        guard let other = PDFDocument(data: file.data), other.pageCount > 0 else { return }
        for index in 0..<min(other.pageCount, 30) {
            if let page = other.page(at: index)?.copy() as? PDFPage {
                document.insert(page, at: document.pageCount)
            }
        }
        return
    }
    guard file.type.hasPrefix("image/"), let image = NSImage(data: file.data), image.size.width >= 32, image.size.height >= 32 else { return }
    addFitted(image, caption: file.name)
}

func addMail(url: URL) {
    guard let data = try? Data(contentsOf: url), !data.isEmpty, let mime = parseMime(data) else {
        addWrapped(title: "E-posta", body: "E-posta metni okunamadı.")
        return
    }
    let piece = gatherMail(mime)
    var lines: [String] = []
    let subject = oneLine(mime.headers["subject"] ?? "", limit: 300)
    let from = oneLine(mime.headers["from"] ?? "", limit: 300)
    let to = oneLine(mime.headers["to"] ?? "", limit: 300)
    let copy = oneLine(mime.headers["cc"] ?? "", limit: 300)
    let date = oneLine(mime.headers["date"] ?? "", limit: 120)
    if !from.isEmpty { lines.append("Kimden: \(from)") }
    if !to.isEmpty { lines.append("Kime: \(to)") }
    if !copy.isEmpty { lines.append("Bilgi: \(copy)") }
    if !date.isEmpty { lines.append("Tarih: \(date)") }
    var body = lines.joined(separator: "\n")
    if !piece.text.isEmpty {
        if !body.isEmpty { body += "\n\n" }
        body += piece.text
    }
    if body.isEmpty { body = "E-posta metni okunamadı." }
    if body.count > 40000 {
        body = String(body.prefix(40000)) + "\n\n… (metin kısaltıldı)"
    }
    addWrapped(title: subject.isEmpty ? "E-posta" : subject, body: body)
    for file in piece.files.prefix(12) {
        addMailFile(file)
    }
}

addWrapped(
    title: job.party.isEmpty ? "Harcama raporu" : job.party,
    body: "\(job.meta)\n\nHarcama nedeni\n\(job.reason)"
)

for path in job.files {
    let url = URL(fileURLWithPath: path)
    let ext = url.pathExtension.lowercased()
    let name = url.lastPathComponent
    if ext == "pdf", let other = PDFDocument(url: url), other.pageCount > 0 {
        for index in 0..<other.pageCount {
            if let page = other.page(at: index)?.copy() as? PDFPage {
                document.insert(page, at: document.pageCount)
            }
        }
    } else if ["png", "jpg", "jpeg", "webp", "tif", "tiff", "heic"].contains(ext), let source = NSImage(contentsOf: url), source.size.width > 0 {
        addFitted(source, caption: name)
    } else if ext == "eml" {
        addMail(url: url)
    } else {
        addWrapped(title: "Eklenemedi", body: name)
    }
}

if document.pageCount == 0 {
    fputs("bos pdf\n", stderr)
    exit(1)
}
let output = URL(fileURLWithPath: job.output)
if !document.write(to: output) {
    fputs("yazilamadi\n", stderr)
    exit(1)
}
