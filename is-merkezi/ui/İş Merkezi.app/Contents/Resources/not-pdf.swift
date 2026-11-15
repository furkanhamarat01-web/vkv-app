import AppKit
import Foundation

struct NoteItem: Decodable {
    let type: String
    let kind: String?
    let x: Double?
    let y: Double?
    let x2: Double?
    let y2: Double?
    let text: String?
    let size: Double?
    let color: String?
    let path: String?
    let w: Double?
    let h: Double?
}

struct NoteBlock: Decodable {
    let type: String
    let text: String?
    let size: Double?
    let color: String?
    let path: String?
    let w: Double?
    let h: Double?
    let kind: String?
    let bold: Bool?
    let italic: Bool?
    let underline: Bool?
    let strike: Bool?
    let highlight: String?
    let align: String?
    let indent: Int?
}

struct NoteJob: Decodable {
    let output: String
    let title: String
    let root: String
    let width: Double
    let height: Double
    let items: [NoteItem]
    let flow: Bool?
    let blocks: [NoteBlock]?
}

func layout(_ blocks: [NoteBlock], pageWidth: CGFloat, root: String, header: CGFloat, draw: Bool) -> CGFloat {
    let left: CGFloat = 24
    let right = pageWidth - 24
    var x = left
    var y: CGFloat = 0
    for block in blocks {
        if block.type == "text" {
            let size = CGFloat(max(12, min(block.size ?? 18, 48)))
            var font = NSFont.systemFont(ofSize: size)
            if block.bold == true && block.italic == true {
                font = NSFontManager.shared.convert(font, toHaveTrait: [.boldFontMask, .italicFontMask])
            } else if block.bold == true {
                font = NSFontManager.shared.convert(font, toHaveTrait: .boldFontMask)
            } else if block.italic == true {
                font = NSFontManager.shared.convert(font, toHaveTrait: .italicFontMask)
            }
            var attrs: [NSAttributedString.Key: Any] = [
                .font: font,
                .foregroundColor: ink(block.color),
            ]
            if block.underline == true {
                attrs[.underlineStyle] = NSUnderlineStyle.single.rawValue
            }
            if block.strike == true {
                attrs[.strikethroughStyle] = NSUnderlineStyle.single.rawValue
            }
            if let mark = block.highlight, !mark.isEmpty {
                attrs[.backgroundColor] = ink(mark)
            }
            let align = block.align ?? "left"
            let gutter = left + CGFloat(max(0, min(block.indent ?? 0, 6))) * 18
            let lines = (block.text ?? "").components(separatedBy: "\n")
            for (index, line) in lines.enumerated() {
                if index > 0 {
                    y += size + 8
                    x = gutter
                } else {
                    x = gutter
                }
                if line.isEmpty {
                    y += size + 8
                    x = gutter
                    continue
                }
                var pending = ""
                let words = line.split(separator: " ", omittingEmptySubsequences: false).map(String.init)
                func place(_ lineText: String) {
                    let width = (lineText as NSString).size(withAttributes: attrs).width
                    var origin = gutter
                    if align == "center" {
                        origin = gutter + max(0, (right - gutter - width) / 2)
                    } else if align == "right" {
                        origin = max(gutter, right - width)
                    }
                    if draw {
                        (lineText as NSString).draw(at: NSPoint(x: origin, y: header + y), withAttributes: attrs)
                    }
                    y += size + 8
                    x = gutter
                }
                for (wordIndex, word) in words.enumerated() {
                    let piece = (wordIndex == 0 || pending.isEmpty ? "" : " ") + word
                    let trial = pending + piece
                    let width = (trial as NSString).size(withAttributes: attrs).width
                    if x + width > right && !pending.isEmpty {
                        place(pending)
                        pending = word
                    } else {
                        pending = trial
                    }
                }
                if !pending.isEmpty {
                    let width = (pending as NSString).size(withAttributes: attrs).width
                    var origin = gutter
                    if align == "center" {
                        origin = gutter + max(0, (right - gutter - width) / 2)
                    } else if align == "right" {
                        origin = max(gutter, right - width)
                    }
                    if draw {
                        (pending as NSString).draw(at: NSPoint(x: origin, y: header + y), withAttributes: attrs)
                    }
                    x = origin + width
                }
            }
        } else if block.type == "shape" || block.type == "image" {
            let shapeWidth = CGFloat(max(20, min(block.w ?? 180, 700)))
            let shapeHeight = CGFloat(max(12, min(block.h ?? 72, 500)))
            if x > left {
                y += 28
            }
            y += 8
            if draw && block.type == "shape" {
                let paint = ink(block.color)
                paint.setStroke()
                let rect = NSRect(x: left, y: header + y, width: shapeWidth, height: shapeHeight)
                let path: NSBezierPath
                if block.kind == "oval" {
                    path = NSBezierPath(ovalIn: rect.insetBy(dx: 2, dy: 2))
                } else if block.kind == "line" {
                    path = NSBezierPath()
                    path.move(to: NSPoint(x: rect.minX, y: rect.midY))
                    path.line(to: NSPoint(x: rect.maxX, y: rect.midY))
                } else {
                    path = NSBezierPath(rect: rect.insetBy(dx: 2, dy: 2))
                }
                path.lineWidth = 2
                path.stroke()
            } else if draw, block.type == "image", let path = block.path, allowedImage(path, root: root),
                      let image = NSImage(contentsOfFile: path) {
                image.draw(in: NSRect(x: left, y: header + y, width: shapeWidth, height: shapeHeight))
            }
            y += shapeHeight + 8
            x = left
        }
    }
    return y + 36
}

func ink(_ hex: String?) -> NSColor {
    let raw = (hex ?? "").trimmingCharacters(in: CharacterSet(charactersIn: "#"))
    guard raw.count == 6, let value = Int(raw, radix: 16) else {
        return NSColor(srgbRed: 0.09, green: 0.125, blue: 0.165, alpha: 1)
    }
    return NSColor(
        srgbRed: CGFloat((value >> 16) & 255) / 255,
        green: CGFloat((value >> 8) & 255) / 255,
        blue: CGFloat(value & 255) / 255,
        alpha: 1
    )
}

func allowedImage(_ path: String, root: String) -> Bool {
    let file = URL(fileURLWithPath: path).standardizedFileURL
    let folder = URL(fileURLWithPath: root).standardizedFileURL
    let prefix = folder.path.hasSuffix("/") ? folder.path : folder.path + "/"
    if !file.path.hasPrefix(prefix) { return false }
    let values = try? file.resourceValues(forKeys: [.isSymbolicLinkKey, .isRegularFileKey])
    return values?.isSymbolicLink == false && values?.isRegularFile == true
}

let arguments = CommandLine.arguments
if arguments.count < 2 {
    fputs("not\n", stderr)
    exit(1)
}

let job: NoteJob
do {
    job = try JSONDecoder().decode(NoteJob.self, from: Data(contentsOf: URL(fileURLWithPath: arguments[1])))
} catch {
    fputs("not\n", stderr)
    exit(1)
}

let pageWidth = CGFloat(max(420, min(job.width, 1400)))
let flowing = job.flow == true
let bodyHeight = flowing
    ? CGFloat(max(280, min(layout(job.blocks ?? [], pageWidth: pageWidth, root: job.root, header: 52, draw: false), 4000)))
    : CGFloat(max(280, min(job.height, 4000)))
let header: CGFloat = 52
let pageHeight = bodyHeight + header
var media = CGRect(x: 0, y: 0, width: pageWidth, height: pageHeight)
let data = NSMutableData()
guard let consumer = CGDataConsumer(data: data as CFMutableData),
      let context = CGContext(consumer: consumer, mediaBox: &media, nil) else {
    fputs("not\n", stderr)
    exit(1)
}

context.beginPDFPage(nil)
let graphics = NSGraphicsContext(cgContext: context, flipped: true)
NSGraphicsContext.saveGraphicsState()
NSGraphicsContext.current = graphics
NSColor.white.setFill()
NSRect(x: 0, y: 0, width: pageWidth, height: pageHeight).fill()
let title = job.title.isEmpty ? "Görev notu" : job.title
(title as NSString).draw(
    in: NSRect(x: 24, y: 16, width: pageWidth - 48, height: 24),
    withAttributes: [
        .font: NSFont.systemFont(ofSize: 16, weight: .semibold),
        .foregroundColor: NSColor(srgbRed: 0.09, green: 0.125, blue: 0.165, alpha: 1),
    ]
)

if flowing {
    _ = layout(job.blocks ?? [], pageWidth: pageWidth, root: job.root, header: header, draw: true)
}

if !flowing {
for item in job.items {
    let paint = ink(item.color)
    let originX = CGFloat(item.x ?? 24)
    let originY = header + CGFloat(item.y ?? 24)
    if item.type == "text" {
        let size = CGFloat(max(12, min(item.size ?? 16, 48)))
        let body = item.text ?? ""
        (body as NSString).draw(
            in: NSRect(x: originX, y: originY, width: pageWidth - originX - 24, height: 220),
            withAttributes: [
                .font: NSFont.systemFont(ofSize: size),
                .foregroundColor: paint,
            ]
        )
    } else if item.type == "shape" {
        let endX = CGFloat(item.x2 ?? originX)
        let endY = header + CGFloat(item.y2 ?? (item.y ?? 24))
        let rect = NSRect(
            x: min(originX, endX),
            y: min(originY, endY),
            width: abs(endX - originX),
            height: abs(endY - originY)
        )
        paint.setStroke()
        let path: NSBezierPath
        if item.kind == "oval" {
            path = NSBezierPath(ovalIn: rect)
        } else if item.kind == "line" {
            path = NSBezierPath()
            path.move(to: NSPoint(x: originX, y: originY))
            path.line(to: NSPoint(x: endX, y: endY))
        } else {
            path = NSBezierPath(rect: rect)
        }
        path.lineWidth = 2
        path.stroke()
    } else if item.type == "image", let path = item.path, allowedImage(path, root: job.root),
              let image = NSImage(contentsOfFile: path) {
        image.draw(in: NSRect(x: originX, y: originY, width: CGFloat(item.w ?? 40), height: CGFloat(item.h ?? 40)))
    }
}
}

NSGraphicsContext.restoreGraphicsState()
context.endPDFPage()
context.closePDF()

do {
    try (data as Data).write(to: URL(fileURLWithPath: job.output))
} catch {
    fputs("not\n", stderr)
    exit(1)
}
