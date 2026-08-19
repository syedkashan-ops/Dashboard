from incremental_merger import merge_daily_sales
from excel_writer import ExcelWriter


def main() -> None:
    try:
        merge_result = merge_daily_sales()
        print(f"{merge_result['message']} Added: {merge_result['added']:,}; duplicates skipped: {merge_result['duplicates']:,}.")
        writer = ExcelWriter()
        writer.create_all()
        writer.save()
    except Exception as exc:
        print(f"\nERROR: {exc}")
        input("Press Enter to close...")
        raise


if __name__ == "__main__":
    main()
