# env_generator.py
from constants import Constants
import sys
import os
import logging
from pathlib import Path
import json

# Global CONSTANTS definitions
PREFIX_DATASET = "dataset"
PREFIX_CALCULATORS = "calculators"
MANIFEST_JSON = "manifest.json"
MISSING_DATASET_CSV = "MISSING_DATASET.csv"

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


def generate_manifest(data_dir, mode):

    manifest_data = {}
    const = Constants(data_dir, mode)

    # build the dataset section of the manifest.json
    def build_dataset_json(path_dataset_csv):
        # build dataset json structure
        json_dataset = {'path': path_dataset_csv.replace('\\', '/')}
        manifest_data['dataset'] = json_dataset
        logging.debug(f"dataset json structure {json.dumps(json_dataset, sort_keys=False, indent=4)}")

    # build the calculators section of the manifest.json
    def build_calculator_json(calculators):
        # build dataset json structure
        json_calculators = []

        for calc in calculators:
            calc_name = Path(calc).name
            logging.debug(f"Calculator instance == {calc}")
            json_calculators.append({'path': calc.replace('\\', '/')})

        logging.debug(f"calculators json structure {json.dumps(json_calculators, sort_keys=False, indent=4)}")
        manifest_data['calculators'] = json_calculators

    manifest_path = Path(data_dir) / MANIFEST_JSON

    # check if we have a manifest.json
    if not manifest_path.exists():
        logging.info(f"No hay un manifest.json en {data_dir}, lo estoy creando {str(manifest_path.resolve())}")

        # check if we have a dataset.csv file
        csv_extensions = ['csv']
        dataset_absolute_path = str(const.get_path_data_set().resolve())
        discovered_datasets = [os.path.join(dataset_absolute_path, fn) for fn in os.listdir(str(const.get_path_data_set().resolve()))
                      if any(fn.endswith(ext) for ext in csv_extensions)]

        # if we do not have a dataset, create a dummy record
        if len(discovered_datasets) < 1:
            logging.warning(f"No hay un dataset en formato csv en {str(const.get_path_data_set().resolve())}")
            dataset_csv_path = const.get_path_data_set() / MISSING_DATASET_CSV
            # build json structure
            build_dataset_json(str(dataset_csv_path.resolve()))

        # if we have datasets, grab the first one
        else:
            logging.info(f"Hay un dataset en formato csv en {str(const.get_path_data_set().resolve())}")
            # sort the list so we take the first csv files in alphabetical order
            discovered_datasets.sort()
            dataset_csv_path = discovered_datasets[0]
            logging.info(f"Utilizando {dataset_csv_path} como el dataset")
            # build json structure
            build_dataset_json(dataset_csv_path)

        # check if we have calculators and add them
        excel_extensions = ['xlsm']
        calculators_absolute_path = str(const.get_path_calculators().resolve())
        discovered_calculators = [os.path.join(calculators_absolute_path, fn) for fn in os.listdir(str(const.get_path_calculators().resolve()))
                      if any(fn.endswith(ext) for ext in excel_extensions)]

        if len(discovered_calculators) < 1:
            logging.error(f"No hay ninguna calculadora en {str(const.get_path_calculators().resolve())}")
            logging.error(f"Tienes que añadir al menos una calculadora en {str(const.get_path_calculators().resolve())}. Exiting.")
            sys.exit
        # if we have calculators then add them to the manifest.json
        else:
            # add the discovered calculators to the json manifest
            build_calculator_json(discovered_calculators)

        # write the gathered data to manifest.json
        json_data = json.dumps(manifest_data, sort_keys=False, indent=4)
        logging.info(f"manifest.json == {json_data}")
        outfile = open(str(manifest_path.resolve()), 'w')
        outfile.write(json_data)

        # exit to give the user the chance to verify the manifest.json file
        logging.info(f"Por favor que controles el manifest.json que ha sido creado en {str(manifest_path.resolve())} ")
        sys.exit(0)
    else:
        logging.info(f"manifest.json ya existe en {manifest_path}")


def generate_directories(data_dir):
    logging.debug(f"data_dir method argument == {data_dir}")
    if os.path.isdir(data_dir):
        logging.info(f"La carpeta proporcionado existe: {data_dir}")
        check_dataset = Path(data_dir) / PREFIX_DATASET
        check_calculators = Path(data_dir) / PREFIX_CALCULATORS

        if check_dataset.exists() and check_dataset.is_dir():
            logging.info(f"La carpeta dataset existe: {check_dataset}")
        else:
            logging.info(f"La carpeta dataset no existe, creando la carpeta: {check_dataset}")
            check_dataset.mkdir(parents=False, exist_ok=False)

        if check_calculators.exists() and check_calculators.is_dir():
            logging.info(f"La carpeta calculators existe: {check_calculators}")
        else:
            logging.info(f"La carpeta calculators no existe, creando la carpeta: {check_calculators}")
            Path(check_calculators).mkdir(parents=False, exist_ok=False)
    else:
        logging.error(f"{data_dir} no existe o no es una carpeta")
        sys.exit
