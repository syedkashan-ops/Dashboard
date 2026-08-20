from incremental_merger import merge_daily_sales
from excel_writer import ExcelWriter


def main() -> None:
    try:
        print("PROGRESS: 1/5 Merging daily sales files...", flush=True)
        merge_result = merge_daily_sales()
        print(
            f"PROGRESS: Daily sales merge completed. Added: {merge_result['added']:,}; "
            f"duplicates skipped: {merge_result['duplicates']:,}.",
            flush=True,
        )

        print("PROGRESS: 2/5 Loading source workbooks and preparing data...", flush=True)
        writer = ExcelWriter()

        print("PROGRESS: 3/5 Creating dashboard and detail sheets...", flush=True)
        writer.create_all()

        print("PROGRESS: 4/5 Saving Result.xlsx...", flush=True)
        writer.save()

        print("PROGRESS: 5/5 Dashboard generation completed successfully.", flush=True)

    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", flush=True)
        raise


if __name__ == "__main__":
    main()
