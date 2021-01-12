# main.py
from constants import Constants
import env_generator as EnvGenerator
import parser_csv as CsvParser
import gcp_sheets
import logging
from pathlib import Path
import argparse
import json
import subprocess
import os
import sys

# Global CONSTANTS definitions
PROGRAM_NAME = "day-offers-dashboard"
MODE_CREATE_MANIFEST_ONLY = "CREATE_MANIFEST_ONLY"
MODE_BUILD_DATASET_ONLY = "BUILD_DATASET_ONLY"
MODE_PUBLISH_DATASET_ONLY = "PUBLISH_DATASET_ONLY"
MODE_CREATE_BUILD_PUBLISH = "CREATE_BUILD_PUBLISH"
MODES = [MODE_CREATE_MANIFEST_ONLY, MODE_BUILD_DATASET_ONLY, MODE_PUBLISH_DATASET_ONLY, MODE_CREATE_BUILD_PUBLISH]
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/"

# logging configuration
if not Path(Constants.get_log_dir()).exists():
    Path(Constants.get_log_dir()).mkdir(parents=False, exist_ok=False)

logging.basicConfig(
    format='%(name)s - %(levelname)s - %(message)s',
    level=logging.DEBUG,
    handlers=[
        logging.FileHandler(Constants.get_log_file()),
        logging.StreamHandler()
    ]
)


def validate_arguments(args):
    if args.data_dir != "":
        if not Path(args.data_dir).exists() or not Path(args.data_dir).is_dir():
            logging.error(f"The provided path argument {args.data_dir} does not exist or is not a directory.")
            sys.exit(99)
    if args.mode != "":
        if args.mode not in MODES:
            logging.error(f"Invalid mode {args.mode} argument has been provided. Valid mode arguments are {str(MODES)}")
            sys.exit(99)
    if args.sheet_id == "":
        logging.error(f"Invalid sheet id {args.sheet_id}. sheet_id cannot be an empty value.")
        sys.exit(99)


def generate_manifest(data_dir, mode):
    const = Constants(data_dir, mode)
    # verify the provided path and generate underlying substructure
    EnvGenerator.generate_directories(data_dir)
    logging.debug(f"path_datadir == {str(const.get_path_data_dir().resolve())}")
    logging.debug(f"path_dataset == {str(const.get_path_data_set().resolve())}")
    logging.debug(f"path_calculators == {str(const.get_path_calculators().resolve())}")

    # with a validate data directory structure, we can build the manifest.json
    EnvGenerator.generate_manifest(data_dir, mode)


def build_dataset(manifest_path):
    # iterate calculators in manifest.json
    with open(manifest_path) as json_file:
        data = json.load(json_file)
        script_path = os.path.dirname(os.path.realpath("main.py")) + "/src"
        row_count_pre = CsvParser.row_count(data['dataset']['path'], True)
        logging.debug(f"Pre update csv row count == {row_count_pre}")
        for calc in data['calculators']:
            # print(f"path: {calc['path']}")
            process = subprocess.run(
                [
                    'C:/Windows/System32/cscript.exe',
                    '/'.join([script_path, 'calculator-to-csv.vbs']),
                    calc['path'],
                    data['dataset']['path']
                ],
                check=True, stdout=subprocess.PIPE, universal_newlines=True)
            output = process.stdout
            logging.debug(output)

        row_count_post = CsvParser.row_count(data['dataset']['path'], True)
        logging.debug(f"Post update csv row count == {row_count_post}")
        row_count_updated = row_count_post - row_count_pre
        if row_count_updated < 1:
            logging.warning(f"{row_count_updated} rows were updated - this indicates a problem")

        logging.info(f"{row_count_updated} rows were updated in: {data['dataset']['path']}")


def main():
    print(str(sys.argv))
    parser = argparse.ArgumentParser(PROGRAM_NAME)
    parser.add_argument("-d", "--data-dir", dest="data_dir", required=True, help="The absolute path to the data directory.")
    parser.add_argument("-m", "--mode", dest="mode", required=True, help="The valid operating modes: " + str(MODES) + ".")
    parser.add_argument("-sid", "--sheet-id", dest="sheet_id", required=True, help="The id of the Google Sheet.")
    args = parser.parse_args()

    # validate the logic of the provided arguments
    validate_arguments(args)
    # Instantiate the Constants class
    const = Constants(args.data_dir, args.mode)

    # generate the environment and manifest in all MODE cases except MODE_PUBLISH_DATASET_ONLY
    if not const.mode == MODE_PUBLISH_DATASET_ONLY:
        # verify the provided path and generate underlying substructure
        generate_manifest(args.data_dir, args.mode)
        # if the mode is MODE_CREATE_MANIFEST_ONLY then exit
        if const.mode == MODE_CREATE_MANIFEST_ONLY:
            logging.info(f"Manifest ha sido creado con exito en {str(const.get_path_manifest_json().resolve())}")
            logging.info(f"Operating mode es {const.mode}. Exiting.")
            sys.exit(0)

    # generate the environment and manifest in all MODE cases except:
    # MODE_CREATE_MANIFEST_ONLY or MODE_PUBLISH_DATASET_ONLY
    if not const.mode == MODE_CREATE_MANIFEST_ONLY and not const.mode == MODE_PUBLISH_DATASET_ONLY:
        build_dataset(str(const.get_path_manifest_json().resolve()))
        if const.mode == MODE_BUILD_DATASET_ONLY:
            logging.info(f"Dataset ha sido creado con exito en {str(const.get_path_data_set().resolve())}")
            logging.info(f"Operating mode es {const.mode}. Exiting.")
            sys.exit(0)

    # publish the dataset if MODE MODE_PUBLISH_DATASET_ONLY or MODE_CREATE_BUILD_PUBLISH
    if const.mode == MODE_PUBLISH_DATASET_ONLY or const.mode == MODE_CREATE_BUILD_PUBLISH:
        with open(str(const.get_path_manifest_json().resolve())) as json_file:
            data = json.load(json_file)
            dataset_csv = data['dataset']['path']
            num_updates = len(data['calculators'])
            logging.info(f"Number of records to update in google sheet: {num_updates}")
            records_to_update = CsvParser.get_records(dataset_csv, num_updates, False)
            for record in records_to_update:
                logging.info(f"Updating record for client: {record['client']}")
                logging.debug(f"Record to update: {record}")

            gcp_sheets.update_workbook(records_to_update, args.sheet_id)

            if const.mode == MODE_PUBLISH_DATASET_ONLY:
                logging.info(f"Dataset ha sido publicado con exito en {GOOGLE_SHEET_URL}{args.sheet_id}")
                logging.info(f"Operating mode es {const.mode}. Exiting.")
                sys.exit(0)

    logging.info("Program exiting correctly.")


if __name__ == "__main__":
    main()
