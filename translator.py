import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from deep_translator import GoogleTranslator
import openpyxl


def get_files_from_directory(path):
    try:
        return os.listdir(path)
    except FileNotFoundError:
        print(f"Directory '{path}' not found.")
        return []


def translate_text(translator, text, translations_cache):
    if text not in translations_cache:
        translations_cache[text] = translator.translate(text)
    return translations_cache[text]


def translate_excel_file(file_path, output_path, translator, translations_cache):
    wb = openpyxl.load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell_value = cell.value.replace(',', '.')
                    try:
                        cell.value = float(cell_value)
                    except ValueError:
                        cell.value = translate_text(translator, cell_value, translations_cache)
    wb.save(output_path)


def process_file(directory, output_directory, file, translator, translations_cache):
    try:
        prefix, name = file.split(') ', 1)
        date = f'{prefix})'
    except ValueError:
        return f"Skipping file: {file} due to invalid format"

    translated_name = translate_text(translator, name, translations_cache)
    output_path = os.path.join(output_directory, f"{date} {translated_name}")

    if os.path.exists(output_path):
        return f"{translated_name} уже переведен"

    if file.endswith(".xlsx"):
        input_path = os.path.join(directory, file)
        translate_excel_file(input_path, output_path, translator, translations_cache)
        return f"Translated and saved: {file} to {output_path}"

    return f"Skipping file: {file} (not an .xlsx file)"


def process_files(directory, output_directory, files):
    translator = GoogleTranslator(source='auto', target='ru')
    translations_cache = {}

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    with ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(process_file, directory, output_directory, file, translator, translations_cache)
            for file in files
        ]

        for future in as_completed(futures):
            print(future.result())


def main():
    print("Введите путь до папки с файлами для перевода:")
    path = input().strip()
    output_directory = os.path.join(path, 'Перевод')

    files = get_files_from_directory(path)
    if files:
        process_files(path, output_directory, files)
    else:
        print("No files found in the directory.")


if __name__ == "__main__":
    main()
