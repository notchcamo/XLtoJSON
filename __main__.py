import sys
from convert import convert_xl_to_json, convert_json_to_xl


def main():
    """Main entry point for XLtoJSON conversion tool.

    Usage:
        python __main__.py <source_file> <output_dir>

    Args:
        source_file: Path to .xlsx or .json file
        output_dir: Directory where output will be saved
    """
    # Validate command line arguments
    if len(sys.argv) < 3:
        print("Usage: python __main__.py <source_file> <output_dir>")
        print("  source_file: Path to .xlsx or .json file")
        print("  output_dir: Directory where output will be saved")
        sys.exit(1)

    source = sys.argv[1]
    dest = sys.argv[2]

    print(f"Source: {source}")
    print(f"Destination: {dest}")

    # Convert based on file extension
    if source.endswith(".xlsx"):
        print("Converting xlsx -> json")
        convert_xl_to_json(source, dest)
        print("Done.")
    elif source.endswith(".json"):
        print("Converting json -> xlsx")
        convert_json_to_xl(source, dest)
        print("Done.")
    else:
        print(f"Error: Unsupported file type. Expected .xlsx or .json, got: {source}")
        sys.exit(1)


if __name__ == "__main__":
    main()
