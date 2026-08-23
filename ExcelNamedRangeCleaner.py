import os
import re
import zipfile
import xml.etree.ElementTree as ET


PROGRAM_NAME = "Excel Named Range Remover"


def get_output_filename(input_file):
    """
    Generate the next revision filename.

    Examples:
        Book.xlsx       -> Book R1.xlsx
        Book R.xlsx     -> Book R1.xlsx
        Book R1.xlsx    -> Book R2.xlsx
        Book r7.xlsx    -> Book R8.xlsx
        Book rx.xlsx    -> Book R1.xlsx
        Book R9.xlsm    -> Book R10.xlsm
    """

    folder = os.path.dirname(input_file)
    filename = os.path.basename(input_file)

    name, extension = os.path.splitext(filename)

    # Detect a trailing:
    #   R
    #   r
    #   R1
    #   r1
    #   R123
    #   r123
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


def remove_all_named_ranges(input_file, output_file):

    # Excel workbook namespace
    MAIN_NS = (
        "http://schemas.openxmlformats.org/"
        "spreadsheetml/2006/main"
    )

    namespace = {
        "main": MAIN_NS
    }

    removed_count = 0

    print("\nOpening Excel file...")

    with zipfile.ZipFile(input_file, "r") as source:

        print(
            f"Excel package contains "
            f"{len(source.infolist()):,} internal files."
        )

        with zipfile.ZipFile(
            output_file,
            "w",
            compression=zipfile.ZIP_DEFLATED
        ) as destination:

            for item in source.infolist():

                # Copy every file exactly as-is except
                # xl/workbook.xml.
                data = source.read(item.filename)

                if item.filename == "xl/workbook.xml":

                    print(
                        "\nFound xl/workbook.xml."
                    )

                    print(
                        "Checking for named ranges..."
                    )

                    root = ET.fromstring(data)

                    defined_names = root.find(
                        "main:definedNames",
                        namespace
                    )

                    if defined_names is None:

                        print(
                            "No named ranges were found."
                        )

                    else:

                        removed_count = len(
                            defined_names
                        )

                        print(
                            f"Found {removed_count:,} "
                            "named range(s)."
                        )

                        print(
                            "Removing the entire "
                            "definedNames section..."
                        )

                        # Remove the entire definedNames
                        # element in one operation.
                        root.remove(
                            defined_names
                        )

                        # Register the original namespace.
                        ET.register_namespace(
                            "",
                            MAIN_NS
                        )

                        data = ET.tostring(
                            root,
                            encoding="utf-8",
                            xml_declaration=True
                        )

                        print(
                            f"Removed {removed_count:,} "
                            "named range(s)."
                        )

                destination.writestr(
                    item,
                    data
                )

    return removed_count


def main():

    print("=" * 72)
    print(PROGRAM_NAME)
    print("=" * 72)

    print(
        "\nThis program removes ALL named ranges "
        "from an Excel workbook."
    )

    print(
        "\nThere are NO exceptions."
    )

    print(
        "Print areas will be removed."
    )

    print(
        "Print titles will be removed."
    )

    print(
        "Every other named range will also be removed."
    )

    print(
        "\nThe original Excel file will NOT be modified."
    )

    print(
        "A new revisioned copy will be created."
    )

    print(
        "\nThe program does not open Excel and does not "
        "use VBA."
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
            "\nERROR: No input file was provided."
        )

        input(
            "\nPress Enter to exit..."
        )

        return

    if not os.path.isfile(input_file):

        print(
            f"\nERROR: File not found:\n"
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

    supported_extensions = (
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm"
    )

    if extension not in supported_extensions:

        print(
            f"\nWARNING: '{extension}' is not a standard "
            "Excel Open XML workbook."
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
    # Generate output filename
    # ---------------------------------------------------------------

    output_file = get_output_filename(
        input_file
    )

    print(
        f"\nInput file:"
        f"\n{input_file}"
    )

    print(
        f"\nOutput file:"
        f"\n{output_file}"
    )

    print("\n" + "-" * 72)

    # ---------------------------------------------------------------
    # Process workbook
    # ---------------------------------------------------------------

    try:

        removed_count = remove_all_named_ranges(
            input_file,
            output_file
        )

        print("\n" + "=" * 72)
        print("PROCESS COMPLETED")
        print("=" * 72)

        print(
            f"\nNamed ranges removed: "
            f"{removed_count:,}"
        )

        print(
            f"\nOutput file:"
            f"\n{output_file}"
        )

        print(
            "\nThe original file was not modified."
        )

        print(
            "\nIMPORTANT:"
        )

        print(
            "Because ALL named ranges were removed, "
            "print areas and print titles have also "
            "been removed."
        )

        print(
            "\nPlease open the output file in Excel "
            "and verify it."
        )

    except zipfile.BadZipFile:

        print(
            "\nERROR: The selected file is not a valid "
            "Excel Open XML workbook."
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
