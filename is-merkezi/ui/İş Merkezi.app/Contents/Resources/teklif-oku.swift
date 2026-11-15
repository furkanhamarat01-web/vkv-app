import AppKit
import Foundation
import PDFKit
import Vision

guard CommandLine.arguments.count > 1 else { exit(1) }
let url = URL(fileURLWithPath: CommandLine.arguments[1])

func recognize(_ image: CGImage) -> String {
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = true
    request.recognitionLanguages = ["tr-TR", "en-US"]
    let handler = VNImageRequestHandler(cgImage: image, options: [:])
    do {
        try handler.perform([request])
    } catch {
        return ""
    }
    return (request.results ?? []).compactMap { $0.topCandidates(1).first?.string }.joined(separator: "\n")
}

func cgImage(from image: NSImage) -> CGImage? {
    var rect = CGRect(origin: .zero, size: image.size)
    return image.cgImage(forProposedRect: &rect, context: nil, hints: nil)
}

func meaningful(_ text: String) -> Bool {
    text.unicodeScalars.filter { CharacterSet.letters.contains($0) }.count >= 40
}

func emit(_ text: String) {
    let clipped = String(text.prefix(20000))
    FileHandle.standardOutput.write(Data(clipped.utf8))
}

let ext = url.pathExtension.lowercased()
if ext == "pdf" {
    guard let document = PDFDocument(url: url) else { exit(1) }
    var parts: [String] = []
    let limit = min(document.pageCount, 5)
    for index in 0..<limit {
        guard let page = document.page(at: index) else { continue }
        let raw = page.string ?? ""
        if meaningful(raw) {
            parts.append(raw)
            continue
        }
        let thumb = page.thumbnail(of: CGSize(width: 1200, height: 1600), for: .mediaBox)
        if let image = cgImage(from: thumb) {
            parts.append(recognize(image))
        }
    }
    emit(parts.joined(separator: "\n"))
    exit(0)
}

if let image = NSImage(contentsOf: url), let cg = cgImage(from: image) {
    emit(recognize(cg))
    exit(0)
}
exit(1)
