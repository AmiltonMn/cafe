import asyncio
import re
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


URL = "https://www.nescafe-dolcegusto.com.br/do-seu-jeito#/capsule-selection/100"
OUTPUT_DIR = Path(__file__).parent
CATEGORIES = ["Lançamentos", "Cafés", "Lattes", "Chocolates", "Chás", "Starbucks"]


async def scrape_capsules():
    products = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        await page.goto(URL, wait_until="networkidle", timeout=60000)

        # Espera os produtos carregarem
        await page.wait_for_selector(".spc-product__name", timeout=30000)

        # Coleta por categoria
        containers = await page.query_selector_all(".spc-listing__container")

        for container in containers:
            cat_name_el = await container.get_attribute("name")
            category = cat_name_el.strip() if cat_name_el else "Outros"

            cards = await container.query_selector_all(".spc-product")
            for card in cards:
                name_el = await card.query_selector(".spc-product__name")
                price_el = await card.query_selector("[data-automation='product-price']")
                pods_el = await card.query_selector(".spc-product__pods-qty")
                sku = await card.get_attribute("data-sku")

                name = (await name_el.inner_text()).strip() if name_el else "—"
                price_text = (await price_el.inner_text()).strip() if price_el else ""
                pods_text = (await pods_el.inner_text()).strip() if pods_el else ""

                # Limpa o preço: "R$2,79" → 2.79
                price_clean = re.sub(r"[^\d,]", "", price_text).replace(",", ".")
                try:
                    price = float(price_clean)
                except ValueError:
                    price = None

                # Extrai quantidade de cápsulas
                pods_match = re.search(r"(\d+)", pods_text)
                pods = int(pods_match.group(1)) if pods_match else None

                products.append({
                    "categoria": category,
                    "nome": name,
                    "preco": price,
                    "capsulas": pods,
                    "sku": sku or "—",
                })

        await browser.close()

    return products


def build_excel(products: list) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cápsulas"

    # Estilos
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="B00000")  # vermelho Dolce Gusto
    cat_font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    cat_fill = PatternFill("solid", start_color="5C0000")
    data_font = Font(name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Título
    ws.merge_cells("A1:F1")
    ws["A1"] = f"Nescafé Dolce Gusto — Preços das Cápsulas"
    ws["A1"].font = Font(name="Arial", bold=True, size=13, color="B00000")
    ws["A1"].alignment = center

    ws.merge_cells("A2:F2")
    ws["A2"] = f"Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = Font(name="Arial", size=9, color="888888")
    ws["A2"].alignment = center

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16

    # Cabeçalho
    headers = ["#", "Categoria", "Nome", "Preço (R$)", "Cápsulas", "SKU"]
    col_widths = [5, 16, 42, 13, 12, 14]

    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[3].height = 20

    # Agrupa por categoria
    from itertools import groupby
    sorted_products = sorted(products, key=lambda x: (CATEGORIES.index(x["categoria"]) if x["categoria"] in CATEGORIES else 99, x["nome"]))

    row = 4
    counter = 1
    alt = False

    for category, items in groupby(sorted_products, key=lambda x: x["categoria"]):
        items = list(items)

        # Linha de categoria
        ws.merge_cells(f"A{row}:F{row}")
        cell = ws.cell(row=row, column=1, value=f"  {category.upper()}  ({len(items)} produtos)")
        cell.font = cat_font
        cell.fill = cat_fill
        cell.alignment = left
        ws.row_dimensions[row].height = 18
        row += 1

        for p in items:
            fill_color = "FFF5F5" if alt else "FFFFFF"
            row_fill = PatternFill("solid", start_color=fill_color)

            values = [counter, p["categoria"], p["nome"], p["preco"], p["capsulas"], p["sku"]]
            aligns = [center, center, left, center, center, center]

            for col, (val, aln) in enumerate(zip(values, aligns), 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = data_font
                cell.fill = row_fill
                cell.alignment = aln
                cell.border = border

                if col == 4 and val is not None:
                    cell.number_format = 'R$#,##0.00'

            ws.row_dimensions[row].height = 17
            counter += 1
            alt = not alt
            row += 1

    # Totais
    ws.merge_cells(f"A{row}:C{row}")
    ws.cell(row=row, column=1, value=f"Total: {len(products)} produtos").font = Font(name="Arial", bold=True, size=10)
    ws.cell(row=row, column=1).alignment = left

    ws.freeze_panes = "A4"

    date_str = datetime.now().strftime("%Y%m%d_%H%M")
    output_path = OUTPUT_DIR / f"capsulas_dolcegusto_{date_str}.xlsx"
    wb.save(output_path)
    return output_path


async def main():
    print("🔍 Acessando site da Dolce Gusto...")
    products = await scrape_capsules()

    if not products:
        print("❌ Nenhum produto encontrado. Verifique o site ou a conexão.")
        return

    print(f"✅ {len(products)} cápsulas encontradas.")
    path = build_excel(products)
    print(f"📊 Planilha salva em: {path}")


if __name__ == "__main__":
    asyncio.run(main())
