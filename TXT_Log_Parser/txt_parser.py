from pathlib import Path
import re

INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
REJECTED_DIR = Path("rejected")

VALID_SYMBOLS = {"S", "R", "E", "N", "X", "Y"}

SYMBOL_PRIORITY = {
    "S": 1,
    "R": 2,
    "E": 3,
    "N": 4,
    "Y": 5,
    "X": 6,
}


def ensure_directories():
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    REJECTED_DIR.mkdir(exist_ok=True)


def sanitize_filename(name):
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", "_", name)
    return name


def is_hex4(value):
    return bool(re.fullmatch(r"[0-9A-Fa-f]{4}", value))


def fix_stuck_symbols(line):

    previous = None

    while previous != line:
        previous = line

        line = re.sub(
            r"([0-9A-Fa-f]{4})([SRNEXY])\b",
            r"\1 \2",
            line,
        )

    return line


def parse_line(line):

    line = fix_stuck_symbols(line)

    tokens = re.split(r"\s+", line.strip())

    if len(tokens) < 4:
        return None

    first_four = tokens[:4]

    if not all(is_hex4(x) for x in first_four):
        return None

    value = ""
    symbols = []

    for token in tokens[4:]:

        if is_hex4(token):

            if value:
                return None

            value = token

        elif token in VALID_SYMBOLS:

            symbols.append(token)

        else:
            return None

    if len(symbols) > 2:
        return None

    symbols.sort(key=lambda x: SYMBOL_PRIORITY[x])

    sym1 = symbols[0] if len(symbols) >= 1 else ""
    sym2 = symbols[1] if len(symbols) >= 2 else ""

    return [
        first_four[0],
        first_four[1],
        first_four[2],
        first_four[3],
        value,
        sym1,
        sym2,
    ]


def process_file(file_path, output_name=None):

    valid_rows = []
    rejected_rows = []

    fixed_symbol_count = 0

    with open(
        file_path,
        "r",
        encoding="utf-8",
        errors="ignore",
    ) as f:
        lines = f.readlines()

    for original_line in lines:

        stripped = original_line.strip()

        if not stripped:
            continue

        if stripped.lower() == "<br>":
            continue

        fixed = fix_stuck_symbols(stripped)

        if fixed != stripped:
            fixed_symbol_count += 1

        parsed = parse_line(fixed)

        if parsed:
            valid_rows.append("\t".join(parsed))
        else:
            rejected_rows.append(stripped)

    base_name = output_name if output_name else file_path.stem

    processed_file = (
        OUTPUT_DIR /
        f"{base_name}_processed.txt"
    )

    with open(
        processed_file,
        "w",
        encoding="utf-8",
    ) as f:
        f.write("\n".join(valid_rows))

    if rejected_rows:

        rejected_file = (
            REJECTED_DIR /
            f"{base_name}_rejected.txt"
        )

        with open(
            rejected_file,
            "w",
            encoding="utf-8",
        ) as f:
            f.write("\n".join(rejected_rows))

    print()
    print(f"Plik               : {file_path.name}")
    print(f"Poprawne rekordy   : {len(valid_rows)}")
    print(f"Odrzucone rekordy  : {len(rejected_rows)}")
    print(f"Naprawione symbole : {fixed_symbol_count}")
    print(f"Wynik              : {processed_file.name}")

    if rejected_rows:
        print(f"Błędy              : {base_name}_rejected.txt")
    else:
        print("Błędy              : brak")


def get_txt_files():
    return sorted(INPUT_DIR.glob("*.txt"))


def ask_naming_mode():

    print("\nNazewnictwo plików:")
    print("1 - Nowa nazwa + numeracja")
    print("2 - Dodaj prefix")
    print("3 - Bez zmian")

    mode = input(
        "\nWybór nazewnictwa: "
    ).strip()

    if mode == "1":

        base_name = sanitize_filename(
            input(
                "Podaj nową nazwę: "
            )
        )

        return 1, base_name, ""

    elif mode == "2":

        prefix = sanitize_filename(
            input(
                "Podaj prefix: "
            )
        )

        return 2, "", prefix

    return 3, "", ""


def process_one(files, naming_mode, base_name, prefix):

    print("\nDostępne pliki:\n")

    for idx, file in enumerate(files, start=1):
        print(f"{idx}. {file.name}")

    try:

        choice = int(
            input("\nWybierz numer pliku: ")
        )

        if not (1 <= choice <= len(files)):
            print("Nieprawidłowy numer.")
            return

        file = files[choice - 1]

        output_name = None

        if naming_mode == 1:
            output_name = f"{base_name}_001"

        elif naming_mode == 2:
            output_name = f"{prefix}_{file.stem}"

        process_file(
            file,
            output_name
        )

    except ValueError:
        print("Podaj liczbę.")


def process_all(files, naming_mode, base_name, prefix):

    for idx, file in enumerate(files, start=1):

        output_name = None

        if naming_mode == 1:

            output_name = (
                f"{base_name}_{idx:03d}"
            )

        elif naming_mode == 2:

            output_name = (
                f"{prefix}_{file.stem}"
            )

        process_file(
            file,
            output_name
        )


def main():

    ensure_directories()

    files = get_txt_files()

    print("=" * 60)
    print("TXT LOG PARSER")
    print("=" * 60)

    if not files:

        print("\nBrak plików TXT w folderze 'input'.")
        print(f"\nŚcieżka:\n{INPUT_DIR.resolve()}")

        input("\nENTER aby zakończyć...")
        return

    print(f"\nWykryto {len(files)} plik(ów):\n")

    for idx, file in enumerate(files, start=1):
        print(f"{idx}. {file.name}")

    naming_mode, base_name, prefix = ask_naming_mode()

    print("\nOpcje:")
    print("1 - Jeden plik")
    print("2 - Wszystkie pliki")
    print("0 - Wyjście")

    choice = input(
        "\nWybór: "
    ).strip()

    if choice == "1":

        process_one(
            files,
            naming_mode,
            base_name,
            prefix
        )

    elif choice == "2":

        process_all(
            files,
            naming_mode,
            base_name,
            prefix
        )

    elif choice == "0":
        return

    else:
        print("Nieznana opcja.")

    input("\nENTER aby zakończyć...")


if __name__ == "__main__":
    main()