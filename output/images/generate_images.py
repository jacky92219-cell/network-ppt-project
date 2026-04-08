"""
使用 Pollinations.ai 免費生成 PPT 插圖（無需 API key）
"""
import urllib.request
import urllib.parse
import time
import os

OUTPUT_DIR = os.path.dirname(__file__)

IMAGES = [
    {
        "filename": "slide03_rf_signal.png",
        "prompt": (
            "Electromagnetic radio waves emanating from a wireless antenna array, "
            "multipath signal propagation through walls, multiple reflective wavefronts "
            "creating interference patterns, luminous silver-gray wave forms, "
            "near-black background, cinematic dramatic lighting, photorealistic, 8K, "
            "dark metallic tech aesthetic, no text, no labels"
        ),
    },
    {
        "filename": "slide04_csi_spectrum.png",
        "prompt": (
            "Abstract visualization of OFDM channel state information, frequency domain "
            "subcarriers displayed as precise silver spectral bars of varying amplitudes, "
            "complex channel matrix floating in deep dark space, metallic scientific "
            "data visualization, dramatic lighting, photorealistic render, 8K, "
            "dark metallic aesthetic, no text, no labels"
        ),
    },
    {
        "filename": "slide08_nic_hardware.png",
        "prompt": (
            "Extreme macro photography of a WiFi 6 M.2 network interface card PCB, "
            "RF shield cans, fine circuit traces, SMD components, antenna connectors, "
            "dramatic silver-gray side lighting, deep black background, "
            "technical product photography, photorealistic, 8K ultra detail, "
            "dark metallic aesthetic, no text, no labels"
        ),
    },
    {
        "filename": "slide15_linux_open.png",
        "prompt": (
            "Abstract visualization of open layered software architecture, "
            "transparent stacked horizontal layers with glowing data streams flowing freely, "
            "contrasting with sealed impenetrable monolithic structure on opposite side, "
            "dark metallic environment, silver highlights, chiaroscuro lighting, "
            "8K cinematic render, photorealistic, no text, no labels"
        ),
    },
    {
        "filename": "slide22_hardware_cards.png",
        "prompt": (
            "Multiple WiFi 6 M.2 network adapter cards arranged on dark reflective surface, "
            "gold edge connectors, PCB with silver components, RF antenna connectors, "
            "dramatic product photography, silver metallic reflections, deep black background, "
            "high-end tech equipment aesthetic, 8K photorealistic, no text, no labels"
        ),
    },
]


def download_image(prompt: str, filename: str, width: int = 1280, height: int = 720) -> bool:
    encoded = urllib.parse.quote(prompt)
    url = (
        f"https://image.pollinations.ai/prompt/{encoded}"
        f"?width={width}&height={height}&model=flux&nologo=true&enhance=true"
    )
    out_path = os.path.join(OUTPUT_DIR, filename)
    try:
        print(f"  Downloading: {filename}")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = resp.read()
        with open(out_path, "wb") as f:
            f.write(data)
        size_kb = len(data) // 1024
        print(f"  Saved: {out_path} ({size_kb} KB)")
        return True
    except Exception as e:
        print(f"  ERROR: {e}")
        return False


if __name__ == "__main__":
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    results = []
    for i, img in enumerate(IMAGES):
        print(f"\n[{i+1}/{len(IMAGES)}] {img['filename']}")
        ok = download_image(img["prompt"], img["filename"])
        results.append((img["filename"], ok))
        if i < len(IMAGES) - 1:
            time.sleep(2)  # 避免過快請求

    print("\n=== 結果 ===")
    for name, ok in results:
        status = "✓" if ok else "✗"
        print(f"  {status} {name}")
