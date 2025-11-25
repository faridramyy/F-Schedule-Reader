import dotenv
from config import TARGET_NAME , HOURLY_RATE, TAX_RATE
from converter import convert_xls_to_xlsx_with_colors
from parser import analyze_schedule
from telegram_downloader import wait_for_xls
import os

def main():
    dotenv.load_dotenv()   
    BOT_TOKEN = os.getenv("BOT_TOKEN")

    print("\n=== TELEGRAM XLS DOWNLOADER ===\n")
    xls_path = wait_for_xls(BOT_TOKEN)
    converted = convert_xls_to_xlsx_with_colors(xls_path)

    if converted:
        analyze_schedule(
            converted,
            TARGET_NAME,
            rate=HOURLY_RATE,
            tax_rate=TAX_RATE
        )

    # Cleanup
    print("\n🧹 Cleaning up temporary files...")

    if converted and os.path.exists(converted):
        try:
            os.remove(converted)
            print(f"Deleted: {converted}")
        except:
            print(f"Could not delete: {converted}")
    if xls_path and os.path.exists(xls_path):
        try:
            os.remove(xls_path)
            print(f"Deleted: {xls_path}")
        except:
            print(f"Could not delete: {xls_path}")

    print("\n✨ Done.\n")


if __name__ == "__main__":
    while True:
        main()