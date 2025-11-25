import os
from converter import convert_xls_to_xlsx_with_colors
from parser import analyze_schedule
from config import FILE_NAME, TARGET_NAME, HOURLY_RATE, TAX_RATE


def main():
    print("\n=== PIZZA HUT SCHEDULE PARSER ===\n")

    converted = convert_xls_to_xlsx_with_colors(FILE_NAME)

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

    print("\n✨ Done.\n")


if __name__ == "__main__":
    main()
