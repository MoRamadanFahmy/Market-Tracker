import requests
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
from datetime import datetime
import pandas as pd
import os
from dotenv import load_dotenv   # to load secrets from .env file

# ===== Load environment variables =====
# Create a .env file in the project root with your API keys and credentials:
# GOLD_API_KEY=*****
# EXCHANGE_API_KEY=*****
# EMAIL_SENDER=*****
# EMAIL_PASSWORD=*****
# EMAIL_RECEIVER=*****
load_dotenv()

# ===== GoldAPI details =====
GOLD_API_KEY = os.getenv("GOLD_API_KEY", "*****")   # fallback shown as stars
GOLD_BASE_URL = "https://www.goldapi.io/api"
gold_headers = {
    "x-access-token": GOLD_API_KEY,
    "Content-Type": "application/json"
}

# ===== ExchangeRate API details =====
EXCHANGE_API_KEY = os.getenv("EXCHANGE_API_KEY", "*****")
EXCHANGE_BASE_URL = f"https://v6.exchangerate-api.com/v6/{EXCHANGE_API_KEY}/latest"

# ===== Email details =====
EMAIL_SENDER = os.getenv("EMAIL_SENDER", "*****")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", "*****")  # Gmail App Password
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER", "*****")

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465  # SSL port

# ===== Fetch Gold Price =====
def fetch_gold_price():
    """Fetch the latest gold price (24k per gram) in EGP."""
    try:
        resp = requests.get(f"{GOLD_BASE_URL}/XAU/EGP", headers=gold_headers, timeout=10)
        data = resp.json()
        return data.get("price_gram_24k")
    except Exception as e:
        print("Error fetching gold price:", e)
        return None

# ===== Fetch Currency & Crypto Rates =====
def fetch_rates():
    """Fetch USD/EGP, EUR/EGP, and BTC & ETH prices in USD."""
    try:
        # USD and EUR to EGP
        resp_usd = requests.get(f"{EXCHANGE_BASE_URL}/USD", timeout=10).json()
        resp_eur = requests.get(f"{EXCHANGE_BASE_URL}/EUR", timeout=10).json()
        usd_to_egp = resp_usd.get("conversion_rates", {}).get("EGP")
        eur_to_egp = resp_eur.get("conversion_rates", {}).get("EGP")

        # Crypto prices in USD from CoinGecko
        resp_crypto = requests.get(
            "https://api.coingecko.com/api/v3/simple/price",
            params={"ids": "bitcoin,ethereum", "vs_currencies": "usd"},
            timeout=10
        ).json()

        btc_to_usd = resp_crypto.get("bitcoin", {}).get("usd")
        eth_to_usd = resp_crypto.get("ethereum", {}).get("usd")

        return usd_to_egp, eur_to_egp, btc_to_usd, eth_to_usd
    except Exception as e:
        print("Error fetching rates:", e)
        return None, None, None, None

# ===== Send Email =====
def send_email(gold_price, usd, eur, btc, eth):
    """Send an email with the latest market update."""
    try:
        subject = "Market Update — Gold, Currency, and Crypto"
        body = (
            f"Gold (24k/gram): {gold_price} EGP\n"
            f"USD → EGP: {usd}\n"
            f"EUR → EGP: {eur}\n"
            f"BTC → USD: {btc}\n"
            f"ETH → USD: {eth}"
        )

        msg = MIMEMultipart()
        msg["From"] = EMAIL_SENDER
        msg["To"] = EMAIL_RECEIVER
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        # Connect to Gmail SMTP server with SSL
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())

        print("✅ Email sent successfully.")

    except Exception as e:
        print("Error sending email:", e)

# ===== Save to Excel =====
def save_to_excel(gold_price, usd, eur, btc, eth):
    """Save market data to Excel file (market_data.xlsx)."""
    file_name = "market_data.xlsx"
    data = {
        "Time": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        "Gold_24k_gram": [gold_price],
        "USD_to_EGP": [usd],
        "EUR_to_EGP": [eur],
        "BTC_to_USD": [btc],
        "ETH_to_USD": [eth]
    }

    df = pd.DataFrame(data)

    # If the file already exists, append new data
    if os.path.exists(file_name):
        df_existing = pd.read_excel(file_name)
        df = pd.concat([df_existing, df], ignore_index=True)

    df.to_excel(file_name, index=False)
    print(f"💾 Data saved to {file_name}")

# ===== Main Loop =====
if __name__ == "__main__":
    while True:
        # Fetch all rates
        gold_price = fetch_gold_price()
        usd, eur, btc, eth = fetch_rates()

        if gold_price and usd and eur and btc and eth:
            # Print to console
            print(f"Gold: {gold_price} EGP | USD: {usd} | EUR: {eur} | BTC: {btc} USD | ETH: {eth} USD")
            
            # Send email and save to Excel
            send_email(gold_price, usd, eur, btc, eth)
            save_to_excel(gold_price, usd, eur, btc, eth)
        else:
            print("⚠️ Failed to fetch some data, skipping...")

        # Wait 30 minutes before running again
        time.sleep(30 * 60)
