import sys
import argparse
from tqdm import tqdm

import ReadData
import App
from libs import Utility, Dialog, Zip, AppConfig

def get_xlsx_paths(inputs):
    files = []
    temp_dirs = []
    
    for input_path in inputs:
        ext = Utility.get_ext_from_path(input_path)
        if ext == "xlsx":
            files.append({"fullpath": input_path, "temp_dir": ""})
        elif ext == "zip":
            extracted_files, temp_dir = Zip.extract_files_from_zip(input_path, extensions=['.xlsx'])
            files.extend([{"fullpath": f, "temp_dir": temp_dir} for f in extracted_files])
            temp_dirs.append(temp_dir)
            
    return files, temp_dirs


def make_selector_label(file, id):
    file_name = file["file"]
    relative_path = f'[{file["relative_path"]}] ' if file["relative_path"] else ""
    return f"{id}: {relative_path}{file_name}"


def file_processor(file, settings, id):
    filename = Utility.get_filename_from_path(filepath=file["fullpath"])
    
    try:
        result = ReadData.aggregate_results(filepath=file["fullpath"], settings=settings)
        if not result or Utility.is_empty(result):
            return {"error": filename}
        
        result["file"] = filename
        result["relative_path"] = (
            Utility.get_relative_directory_path(full_path=file["fullpath"], base_dir=file["temp_dir"])
            if file["temp_dir"] else ""
        )
        result["selector_label"] = make_selector_label(result, id)
        
        return result
    except Exception as e:
        return {"error": f"{filename}: {str(e)}"}


def start():
    parser = argparse.ArgumentParser(description="zipファイル/xlsxファイルを引数として起動します。(複数可)")
    parser.add_argument("--debug", action="store_true", help="デバッグモードを有効化")
    parser.add_argument("data_files", nargs="*", help="zipファイル/xlsxファイルのパス")
    args = parser.parse_args()

    inputs = args.data_files or Dialog.select_files(("Excel/Zipファイル", "*.xlsx;*.zip"))
    if not inputs:
        sys.exit()

    files, temp_dirs = get_xlsx_paths(inputs)
    settings = AppConfig.load_settings()

    results = [
        file_processor(file, settings, i+1)
        for i, file in enumerate(tqdm(files))
    ]
    
    out_data = [r for r in results if "error" not in r]
    errors = [r for r in results if "error" in r]

    if args.debug:
        from pprint import pprint
        pprint(out_data)

    App.launch(out_data, errors, inputs)

    if temp_dirs:
        Zip.cleanup_old_temp_dirs()


if __name__ == "__main__":
    start()