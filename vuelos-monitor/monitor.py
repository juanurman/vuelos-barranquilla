import asyncio
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from playwright.async_api import async_playwright

ALERT_THRESHOLD = 500  # Objetivo: menor a 500 USD

ROUTES = [
    {
        "key": "RIO_30_02",
        "name": "BUE -> RIO (30 Dic - 02 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-30%20through%202027-01-02&curr=USD",
    },
    {
        "key": "RIO_30_03",
        "name": "BUE -> RIO (30 Dic - 03 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-30%20through%202027-01-03&curr=USD",
    },
    {
        "key": "RIO_31_02",
        "name": "BUE -> RIO (31 Dic - 02 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-31%20through%202027-01-02&curr=USD",
    },
    {
        "key": "RIO_31_03",
        "name": "BUE -> RIO (31 Dic - 03 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-31%20through%202027-01-03&curr=USD",
    },
    {
        "key": "IDA_30",
        "name": "Solo Ida BUE -> RIO (30 Dic)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-30&curr=USD",
    },
    {
        "key": "IDA_31",
        "name": "Solo Ida BUE -> RIO (31 Dic)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20RIO%20from%20BUE%20on%202026-12-31&curr=USD",
    },
    {
        "key": "VTA_02",
        "name": "Solo Ida RIO -> BUE (02 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BUE%20from%20RIO%20on%202027-01-02&curr=USD",
    },
    {
        "key": "VTA_03",
        "name": "Solo Ida RIO -> BUE (03 Ene)",
        "url": "https://www.google.com/travel/flights?q=Flights%20to%20BUE%20from%20RIO%20on%202027-01-03&curr=USD",
    },
]


async def scrape_single_route(context, route_info):
  page = await context.new_page()
  try:
    await page.goto(
        route_info["url"], wait_until="domcontentloaded", timeout=35000
    )
    await page.wait_for_timeout(2000)

    btn = page.get_by_role(
        "button",
        name=re.compile(r"Proceed anyway|Continuar|Aceptar|Accept all", re.I),
    )
    if await btn.count() > 0:
      try:
        await btn.first.click()
        await page.wait_for_timeout(2000)
      except Exception:
        pass

    text = await page.inner_text("body")
    m = re.search(r"(?:from|desde)\s*(?:USD|US\$|\$)\s*([\d,]+)", text, re.I)
    if m:
      val = int(m.group(1).replace(",", ""))
      await page.close()
      return val

    matches = [
        int(x.replace(",", ""))
        for x in re.findall(r"(?:USD|US\$|\$)\s*([\d,]+)", text)
        if 50 <= int(x.replace(",", "")) <= 10000
    ]
    await page.close()
    return min(matches) if matches else "Sin dato"
  except Exception as e:
    await page.close()
    return "Error"


async def main():
  async with async_playwright() as p:
    browser = await p.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
        ],
    )
    context = await browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            " (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        ),
        ignore_https_errors=True,
    )

    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Iniciando...")
    alerts = []
    prices = {}

    for r in ROUTES:
      price = await scrape_single_route(context, r)
      prices[r["key"]] = price
      print(f"  {r['name']}: {price}")
      if isinstance(price, (int, float)) and price < ALERT_THRESHOLD:
        alerts.append((r["name"], price))

    # Chequeo de combos Solo Ida + Solo Ida
    for ida_k, ida_nom in [
        ("IDA_30", "Ida 30 Dic"),
        ("IDA_31", "Ida 31 Dic"),
    ]:
      for vta_k, vta_nom in [
          ("VTA_02", "Vuelta 02 Ene"),
          ("VTA_03", "Vuelta 03 Ene"),
      ]:
        if isinstance(prices.get(ida_k), (int, float)) and isinstance(
            prices.get(vta_k), (int, float)
        ):
          combo = prices[ida_k] + prices[vta_k]
          if combo < ALERT_THRESHOLD:
            alerts.append(
                (f"🔥 Combo Desdoblado ({ida_nom} + {vta_nom})", combo)
            )

    await browser.close()

    if alerts:
      print(f"\n🚨 ¡ALERTA! Se encontraron {len(alerts)} opciones < US$ 500:")
      for nom, p_val in alerts:
        print(f"  👉 {nom}: US$ {p_val}")


if __name__ == "__main__":
  asyncio.run(main())
