import os

from docx import Document


def extract_images_from_docx(docx_path: str):
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"File not found: {docx_path}")

    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    output_dir = f"{base_name}_images"
    os.makedirs(output_dir, exist_ok=True)

    doc = Document(docx_path)
    package = doc.part.package

    count = 0
    for part in package.parts:
        if part.content_type.startswith("image/"):
            blob = getattr(part, "blob", None) or getattr(part, "_blob", None)
            if not blob:
                continue

            content_type = part.content_type
            ext = content_type.split("/")[-1]
            if ext == "jpeg":
                ext = "jpg"

            count += 1
            filename = f"image_{count:03d}.{ext}"
            output_path = os.path.join(output_dir, filename)
            with open(output_path, "wb") as f:
                f.write(blob)

            print(f"[{count}] Saved: {output_path} ({len(blob)} bytes)")

    if count == 0:
        print("There aren't images")
    else:
        print(f"Extracted {count} images to {output_dir}")


if __name__ == "__main__":
    file_path = "your.docx"
    extract_images_from_docx(file_path)
