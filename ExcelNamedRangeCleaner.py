import os
import re
import zipfile
import xml.etree.ElementTree as ET


PROGRAM_NAME = "Excel Named Range Cleaner"


def get_next_revision_filename(input_file):
    """
    Creates the next revision filename.

    Examples:

        Book.xlsx       -> Book R1.xlsx
        Book R.xlsx     -> Book R1.xlsx
        Book R1.xlsx    -> Book R2.xlsx
        Book r7.xlsx    -> Book R8.xlsx
        Book Rx.xlsx    -> Book R1.xlsx
        Book rx.xlsx    -> Book R1.xlsx
    """

    folder = os.path.dirname(input_file)
    filename = os.path.basename(input_file)

    name, extension = os.path.splitext(filename)

    # Detect:
    #   " R"
    #   " r"
    #   " R1"
    #   " r1"
    #   " Rx"
    #   " rx"
    #
    # A revision suffix must be at the END of the filename.
    match = re.search(r"\s[Rr](\d*)$", name)

    if match:
        revision_text = match.group(1)

        if revision_text:
            revision = int(revision_text) + 1
        else:
            revision = 1

        base_name = name[:match.start()]

    else:
        revision = 1
        base_name = name

    while True:

        output_name = (
            f"{base_name} R{revision}{extension}"
        )

        output_file = os.path.join(
            folder,
            output_name
        )

        if not os.path.exists(output_file):
            return output_file

        revision += 1


def is_print_name(name):
    """
    Returns True only for:

        _xlnm.Print_Area
        _xlnm.Print_Titles

    The _xlnm prefix is optional because some workbooks
    may represent the names differently.
    """

    name = name.strip().lower()

    if name.startswith("_xlnm."):
        name = name[6:]

    return name in (
        "print_area",
        "print_titles"
    )


def process_workbook(input_file, output_file):

    MAIN_NS = (
        "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main"
    )

    namespace = {
        "main": MAIN_NS
    }

    removed_count = 0
    retained_count = 0

    print("\nOpening Excel package...")

    with zipfile.ZipFile(input_file, "r") as source:

        with zipfile.ZipFile(
            output_file,
            "w",
            compression=zipfile.ZIP_DEFLATED
        ) as destination:

            for item in source.infolist():

                data = source.read(item.filename)

                # Only modify workbook.xml.
                if item.filename == "xl/workbook.xml":

                    print("Processing xl/workbook.xml...")
                    print("Searching for defined names...")

                    root = ET.fromstring(data)

                    defined_names = root.find(
                        "main:definedNames",
                        namespace
                    )

                    if defined_names is None:

                        print(
                            "No defined names were found."
                        )

                    else:

                        original_count = len(
                            defined_names
                        )

                        print(
                            f"Found {original_count:,} "
                            "defined name(s)."
                        )

                        names_to_remove = []

                        for defined_name in defined_names:

                            name = defined_name.get(
                                "name",
                                ""
                            )

                            if is_print_name(name):

                                retained_count += 1

                                print(
                                    f"  KEEP   {name}"
                                )

                            else:

                                names_to_remove.append(
                                    defined_name
                                )

                        print(
                            "\nRemoving unwanted "
                            "defined names..."
                        )

                        for defined_name in names_to_remove:

                            name = defined_name.get(
                                "name",
                                ""
                            )

                            defined_names.remove(
                                defined_name
                            )

                            removed_count += 1

                        # If nothing remains, remove the
                        # definedNames container.
                        if len(defined_names) == 0:

                            root.remove(
                                defined_names
                            )

                        # Register the namespace so that
                        # Excel's namespace remains normal.
                        ET.register_namespace(
                            "",
                            MAIN_NS
                        )

                        data = ET.tostring(
                            root,
                            encoding="utf-8",
                            xml_declaration=True
                        )

                destination.writestr(
                    item,
                    data
                )

    return removed_count, retained_count


def main():

    print("=" * 72)
    print(PROGRAM_NAME)
    print("=" * 72)

    print(
        "\nWHAT THIS PROGRAM DOES"
    )

    print(
        "This program removes Excel defined names "
        "from a workbook."
    )

    print(
        "It keeps ONLY:"
    )

    print(
        "  1. Print_Area"
    )

    print(
        "  2. Print_Titles"
    )

    print(
        "Everything else is removed."
    )

    print(
        "\nThe original workbook is never modified."
    )

    print(
        "A new revisioned workbook is created."
    )

    print(
        "\nThe program works directly on the Excel "
        "file package, so it does not need Excel "
        "or VBA."
    )

    print(
        "\nThis is particularly useful for workbooks "
        "with 100,000+ defined names."
    )

    print("\n" + "-" * 72)

    # ---------------------------------------------------------------
    # Ask for input file
    # ---------------------------------------------------------------

    input_file = input(
        "\nEnter the full path of the Excel file:\n> "
    ).strip().strip('"')

    if not input_file:

        print(
            "\nERROR: No input file was supplied."
        )

        input(
            "\nPress Enter to exit..."
        )

        return

    if not os.path.isfile(input_file):

        print(
            f"\nERROR: File does not exist:\n"
            f"{input_file}"
        )

        input(
            "\nPress Enter to exit..."
        )

        return

    # ---------------------------------------------------------------
    # Check extension
    # ---------------------------------------------------------------

    extension = os.path.splitext(
        input_file
    )[1].lower()

    supported = (
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm"
    )

    if extension not in supported:

        print(
            f"\nWARNING: {extension} is not a supported "
            "Excel Open XML extension."
        )

        answer = input(
            "Continue anyway? (Y/N): "
        ).strip().lower()

        if answer != "y":

            print(
                "\nOperation cancelled."
            )

            input(
                "\nPress Enter to exit..."
            )

            return

    # ---------------------------------------------------------------
    # Determine output
    # ---------------------------------------------------------------

    output_file = get_next_revision_filename(
        input_file
    )

    print(
        f"\nINPUT:"
        f"\n{input_file}"
    )

    print(
        f"\nOUTPUT:"
        f"\n{output_file}"
    )

    print("\n" + "-" * 72)

    # ---------------------------------------------------------------
    # Process
    # ---------------------------------------------------------------

    try:

        removed, retained = process_workbook(
            input_file,
            output_file
        )

        print("\n" + "=" * 72)
        print("PROCESS COMPLETED SUCCESSFULLY")
        print("=" * 72)

        print(
            f"\nDefined names removed : {removed:,}"
        )

        print(
            f"Print names retained  : {retained:,}"
        )

        print(
            f"\nOutput file:"
            f"\n{output_file}"
        )

        print(
            "\nThe original file was NOT modified."
        )

        print(
            "\nYou should open the output workbook "
            "in Excel and verify it before replacing "
            "the original."
        )

    except zipfile.BadZipFile:

        print(
            "\nERROR: The selected file is not a "
            "valid Excel Open XML workbook."
        )

        if os.path.exists(output_file):

            try:
                os.remove(output_file)
            except Exception:
                pass

    except PermissionError:

        print(
            "\nERROR: Windows denied access to the file."
        )

        print(
            "Make sure the workbook is not open in Excel "
            "and that you have permission to write to "
            "the folder."
        )

        if os.path.exists(output_file):

            try:
                os.remove(output_file)
            except Exception:
                pass

    except Exception as error:

        print(
            "\nERROR:"
        )

        print(
            str(error)
        )

        if os.path.exists(output_file):

            try:
                os.remove(output_file)
            except Exception:
                pass

    print("\n" + "-" * 72)

    input(
        "Press Enter to exit..."
    )


if __name__ == "__main__":
    main()