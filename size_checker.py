from docx import Document

doc = Document("your.docx")
package = doc.part.package

for i, part in enumerate(package.parts, start=1):
    if part.content_type.startswith("image/"):
        blob = getattr(part, "blob", None) or getattr(part, "_blob", None)
        if blob:
            size_bytes = len(blob)
            print(f"[{i}] {part.partname}: {size_bytes} bytes")
        else:
            print(f"[{i}] {part.partname}: <no data>")
