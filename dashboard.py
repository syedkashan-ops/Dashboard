from incremental_merger import merge_daily_sales
from excel_writer import ExcelWriter


def main() -> None:
    try:
        print("Starting daily sales merge...", flush=True)

        merge_result = merge_daily_sales()

        print(
            f"{merge_result['message']} "
            f"Added: {merge_result['added']:,}; "
            f"duplicates skipped: {merge_result['duplicates']:,}.",
            flush=True
        )

        print("Creating Excel dashboard...", flush=True)

        writer = ExcelWriter()

        print("Generating dashboard sheets...", flush=True)

        writer.create_all()

        print("Saving Result.xlsx...", flush=True)

        writer.save()

        print(
            "Dashboard generation completed successfully.",
            flush=True
        )

    except Exception as exc:

        print(
            f"\nERROR: {type(exc).__name__}: {exc}",
            flush=True
        )

        # IMPORTANT:
        # Do not use input() here.
        # Streamlit Cloud has no interactive terminal.

        raise


if __name__ == "__main__":
    main()
