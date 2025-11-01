import argparse
import io
import os
import shutil
from typing import Optional

import torch
from PIL import Image

try:
    from docx import Document
except Exception as e:
    Document = None


def _check_imports():
    if Document is None:
        raise RuntimeError("python-docx is required. Install with: pip install python-docx")


def _load_realesrgan_model(device="cpu"):
    try:
        from realesrgan import RealESRGAN
        model = RealESRGAN(torch.device(device), scale=4)
        model.load_weights(RealESRGAN.download_weights("RealESRGAN_x4plus"))
        return model
    except Exception:
        return None


def upscale_pillow(img: Image.Image, scale: float) -> Image.Image:
    w, h = img.size
    new_size = (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
    return img.resize(new_size, Image.LANCZOS)


def upscale_realesrgan(img: Image.Image, scale: float, model=None) -> Optional[Image.Image]:
    if model is None:
        device = "cuda" if torch.cuda.is_available() else "cpu"
        model = _load_realesrgan_model(device=device)
        if model is None:
            return None
    img_rgb = img.convert("RGB")
    sr = model.predict(img_rgb)
    if scale != 4.0:
        w, h = img.size
        target = (max(1, int(round(w * scale))), max(1, int(round(h * scale))))
        sr = sr.resize(target, Image.LANCZOS)
    return sr


def upscale_ncnn_cli(img_bytes: bytes, scale: float) -> Optional[bytes]:
    import subprocess
    import tempfile

    exe = shutil.which("realesrgan-ncnn-vulkan")
    if exe is None:
        return None
    with tempfile.TemporaryDirectory() as tmpd:
        in_path = os.path.join(tmpd, "in.png")
        out_path = os.path.join(tmpd, "out.png")
        try:
            Image.open(io.BytesIO(img_bytes)).save(in_path, format="PNG")
        except Exception:
            with open(in_path, "wb") as f:
                f.write(img_bytes)

        ncnn_scale = int(round(scale))
        if ncnn_scale not in (2, 3, 4):
            ncnn_scale = 4
        cmd = [exe, "-i", in_path, "-o", out_path, "-s", str(ncnn_scale)]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            with open(out_path, "rb") as f:
                return f.read()
        except Exception:
            return None


def upscale_image_bytes(img_bytes: bytes, orig_format: str, scale: float, method: str, jpeg_quality: int,
                        keep_format: bool) -> Optional[bytes]:
    methods = []
    if method == "realesrgan":
        methods = ["realesrgan", "ncnn", "pillow"]
    elif method == "ncnn":
        methods = ["ncnn", "pillow"]
    else:
        methods = ["pillow"]

    for m in methods:
        try:
            if m == "ncnn":
                out = upscale_ncnn_cli(img_bytes, scale)
                if out:
                    return out
            elif m == "realesrgan":
                with Image.open(io.BytesIO(img_bytes)) as im:
                    up = upscale_realesrgan(im, scale)
                    if up is None:
                        continue
                    buf = io.BytesIO()
                    fmt = (im.format or orig_format or "PNG") if keep_format else "PNG"
                    save_kwargs = {}
                    if fmt.upper() in ("JPG", "JPEG"):
                        fmt = "JPEG"
                        save_kwargs["quality"] = jpeg_quality
                    up.save(buf, format=fmt, **save_kwargs)
                    return buf.getvalue()
            else:
                with Image.open(io.BytesIO(img_bytes)) as im:
                    up = upscale_pillow(im, scale)
                    buf = io.BytesIO()
                    fmt = (im.format or orig_format or "PNG") if keep_format else "PNG"
                    save_kwargs = {}
                    if fmt and fmt.upper() in ("JPG", "JPEG"):
                        fmt = "JPEG"
                        save_kwargs["quality"] = jpeg_quality
                    up.save(buf, format=fmt, **save_kwargs)
                    return buf.getvalue()
        except Exception:
            continue
    return None


def process_docx_replace(input_docx: str, scale: float = 2.0, method: str = "pillow", jpeg_quality: int = 95,
                         keep_format: bool = True, output_docx: Optional[str] = None) -> str:
    _check_imports()
    if not input_docx.lower().endswith(".docx"):
        raise ValueError("Only .docx files are supported.")
    if output_docx is None:
        base, _ = os.path.splitext(input_docx)
        output_docx = f"{base}_upscaled.docx"

    doc = Document(input_docx)
    package = doc.part.package

    processed = 0
    skipped = 0

    for part in package.parts:
        ctype = getattr(part, "content_type", "")
        if not ctype.startswith("image/"):
            continue
        blob = getattr(part, "blob", None)
        if blob is None:
            blob = getattr(part, "_blob", None)
        if not blob:
            skipped += 1
            continue

        orig_fmt = None
        if ctype.startswith("image/"):
            orig_fmt = ctype.split("/", 1)[1].upper()
            if orig_fmt == "JPG":
                orig_fmt = "JPEG"

        new_bytes = upscale_image_bytes(blob, orig_fmt, scale, method, jpeg_quality, keep_format)
        if new_bytes:
            try:
                part._blob = new_bytes
                processed += 1
            except Exception:
                skipped += 1
        else:
            skipped += 1

    doc.save(output_docx)
    return output_docx


def cli():
    parser = argparse.ArgumentParser(
        description="Upscale and REPLACE embedded images inside a .docx using python-docx.")
    parser.add_argument("input", help="Path to input .docx")
    parser.add_argument("--scale", type=float, default=2.0, help="Upscale factor (2, 3, 4...)")
    parser.add_argument("--method", choices=["pillow", "realesrgan", "ncnn"], default="pillow",
                        help="Upscaling backend")
    parser.add_argument("--jpeg-quality", type=int, default=95, help="JPEG quality for saving (if JPEG)")
    parser.add_argument("--convert-to-png", action="store_true", help="Force saving all images as PNG inside DOCX")
    args = parser.parse_args()

    out = process_docx_replace(
        args.input,
        scale=args.scale,
        method=args.method,
        jpeg_quality=args.jpeg_quality,
        keep_format=not args.convert_to_png,
    )
    print(f"Done. Wrote: {out}")


if __name__ == "__main__":
    cli()
