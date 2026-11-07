# -*- coding: utf-8 -*-
"""collect_evidence.py — сбор доказательств копирования контента.

Для каждой карточки-нарушителя (WB артикул/URL, Ozon URL):
  - название, продавец, ID продавца, дата создания карточки (WB)
  - все изображения карточки → evidence/<platform>_<sku>/
  - скриншот страницы карточки
Итог: evidence/evidence_report.xlsx — таблица «нарушитель → материалы».

Запуск:  python collect_evidence.py refs.txt
         refs.txt — по одному артикулу WB или URL (WB/Ozon) на строку.
"""
import io
import os
import re
import sys
import time
import datetime as dt

sys.path.insert(0, os.path.dirname(__file__))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import requests
import openpyxl
from openpyxl.styles import Font


def _basket_num(vol):
    for top, n in [(143, 1), (287, 2), (431, 3), (719, 4), (1007, 5),
                   (1061, 6), (1115, 7), (1169, 8), (1313, 9), (1601, 10),
                   (1655, 11), (1919, 12), (2045, 13), (2337, 14),
                   (2387, 15), (2745, 16), (2819, 17), (2939, 18),
                   (3059, 19), (3179, 20), (3299, 21), (3419, 22)]:
        if vol <= top:
            return f"{n:02d}"
    return "23"

BASE = os.path.dirname(os.path.abspath(__file__))
EV = os.path.join(BASE, "evidence")
os.makedirs(EV, exist_ok=True)

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "Chrome/120.0 Safari/537.36"}


def wb_card(nm):
    """WB по артикулу: название, продавец, seller_id, дата создания, фото."""
    nm = int(nm)
    vol, part = nm // 100000, nm // 1000
    b = _basket_num(vol)
    base = f"https://basket-{b}.wbbasket.ru/vol{vol}/part{part}/{nm}/info"
    out = {"nm": nm, "name": None, "brand": None, "seller": None,
           "seller_id": None, "created": None, "npics": 0}
    try:
        d = None
        for _try in range(5):
            try:
                d = requests.get(f"{base}/ru/card.json", headers=UA, timeout=25).json()
                break
            except Exception as e:
                if _try == 4:
                    raise
                time.sleep(3)
        out["name"] = d.get("imt_name")
        out["brand"] = (d.get("selling") or {}).get("brand_name") or d.get("brand")
        out["created"] = d.get("create_date")
        out["npics"] = (d.get("media") or {}).get("photo_count", 0)
    except Exception as e:
        out["error"] = str(e)[:120]
        return out
    # supplier_id живёт в card.json -> selling.supplier_id;
    # имя продавца — в suppliers-api (если отвечает)
    try:
        out["seller_id"] = (d.get("selling") or {}).get("supplier_id")
        out["brand"] = out["brand"] or (d.get("selling") or {}).get("brand_name")
    except Exception:
        pass
    if out.get("seller_id"):
        try:
            for _try in range(3):
                try:
                    sd = requests.get(
                        f"https://suppliers-api.wildberries.ru/public/v1/sellers/{out['seller_id']}",
                        headers=UA, timeout=15).json()
                    out["seller"] = sd.get("trademark") or sd.get("name")
                    break
                except Exception:
                    if _try < 2:
                        time.sleep(2)
        except Exception:
            pass
    out["pics"] = [
        f"https://basket-{b}.wbbasket.ru/vol{vol}/part{part}/{nm}/images/big/{j}.webp"
        for j in range(1, out["npics"] + 1)]
    out["url"] = f"https://www.wildberries.ru/catalog/{nm}/detail.aspx"
    return out


def save_imgs(urls, folder, prefix="img"):
    saved = []
    for i, u in enumerate(urls, 1):
        try:
            ext = ".webp" if "webp" in u else ".jpg"
            fn = os.path.join(folder, f"{prefix}_{i:02d}{ext}")
            b = requests.get(u, headers=UA, timeout=20).content
            if len(b) > 1000:
                open(fn, "wb").write(b)
                saved.append(fn)
        except Exception:
            pass
    return saved


def screenshot(url, out):
    """Скриншот карточки через main-Chrome (9222)."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.connect_over_cdp("http://127.0.0.1:9222",
                                            timeout=25000)
            page = b.contexts[0].new_page()
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=25000)
            except Exception:
                pass
            time.sleep(4)
            page.screenshot(path=out, full_page=False)
            page.close()
            return True
    except Exception as e:
        print("  screenshot err:", str(e)[:80])
        return False


def main(refs_file):
    no_shot = '--no-shot' in sys.argv
    refs = [l.strip() for l in open(refs_file, encoding="utf-8") if l.strip()]
    rows = []
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    for ref in refs:
        if "ozon" in ref:
            m = re.search(r"-(\d{6,})/?$", ref) or re.search(r"/(\d{6,})/?$", ref)
            sku = m.group(1) if m else re.sub(r"\D", "", ref)
            print("Ozon", sku)
            folder = os.path.join(EV, f"ozon_{sku}")
            os.makedirs(folder, exist_ok=True)
            shot = os.path.join(folder, "screenshot.png")
            ok = False if no_shot else screenshot(ref, shot)
            rows.append({
                "platform": "Ozon", "nm": sku, "name": "(см. скриншот)",
                "seller": "", "seller_id": "", "price": "", "created": "",
                "imgs": 0, "shot": ok, "url": ref,
                "note": "данные карточки — со скриншота",
            })
        else:
            m = re.search(r"catalog/(\d+)", ref)
            digits = re.sub(r"\D", "", m.group(1) if m else ref)
            if not digits:
                rows.append({"platform": "?", "ref": ref,
                             "note": "не распознано"})
                continue
            nm = int(digits)
            print("WB", nm)
            info = wb_card(nm)
            if "error" in info:
                rows.append({"platform": "WB", "nm": nm,
                             "note": info["error"], "url": ref})
                continue
            folder = os.path.join(EV, f"wb_{nm}")
            os.makedirs(folder, exist_ok=True)
            imgs = save_imgs(info["pics"], folder)
            shot = os.path.join(folder, "screenshot.png")
            ok = False if no_shot else screenshot(info["url"], shot)
            rows.append({
                "platform": "WB", "nm": nm, "name": info["name"],
                "seller": info["seller"], "seller_id": info["seller_id"],
                "price": "", "created": info["created"],
                "imgs": len(imgs), "shot": ok,
                "url": info["url"], "note": "",
            })
        time.sleep(1)

    wb_x = openpyxl.Workbook()
    ws = wb_x.active
    ws.title = "evidence"
    ws.append(["Платформа", "SKU/артикул", "Название", "Продавец",
               "ID продавца", "Карточка создана", "Фото скачано", "Скриншот",
               "URL", "Зафиксировано", "Примечание"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([r.get("platform"), r.get("nm"), r.get("name"),
                   r.get("seller"), r.get("seller_id"), r.get("created"),
                   r.get("imgs"), "да" if r.get("shot") else "нет",
                   r.get("url"), stamp, r.get("note")])
    for col, w in zip("ABCDEFGHIJK",
                      (10, 12, 45, 22, 12, 20, 12, 10, 55, 16, 25)):
        ws.column_dimensions[col].width = w
    out = os.path.join(EV, "evidence_report.xlsx")
    wb_x.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "refs.txt")
