# test_main.py
# to get the import statements on the test working, make sure to set PYTHONPATH which includes the src directory
# export PYTHONPATH=/opt/devel/source/dashboard-flow/src
import env_generator as EnvGenerator
import os
import shutil
from pathlib import Path
import json
import pytest

# tests resources definitions
MODE_CREATE_MANIFEST_ONLY = "CREATE_MANIFEST_ONLY"
TEST_RESOURCES_DIR = "resources"
TEST_DATA_DIR = "test_datadir"
TEST_DATASET_DIR = "dataset"
TEST_CALCULATORS_DIR = "calculators"
TEST_MANIFEST_JSON = "manifest.json"
MATCH_PATH_DATASET = str(
    (Path(TEST_DATA_DIR) / TEST_DATASET_DIR / "dataset.csv")
    .absolute()
    .resolve()).replace('\\', '/')
MATCH_PATH_CALCULATOR_1 = str(
    (Path(TEST_DATA_DIR) / TEST_CALCULATORS_DIR / "Donald Duck - Day Pre-Estimation Calculator.xlsm")
    .absolute()
    .resolve()).replace('\\', '/')
MATCH_PATH_CALCULATOR_2 = str(
    (Path(TEST_DATA_DIR) / TEST_CALCULATORS_DIR / "Mickey Mouse - Day Pre-Estimation Calculator.xlsm")
    .absolute()
    .resolve()).replace('\\', '/')


def prepare_test_env():
    if not Path(TEST_DATA_DIR).exists():
        # create directory
        Path(TEST_DATA_DIR).mkdir(parents=False, exist_ok=False)
        # copy directories from resources folder
        path_src_dataset = Path(TEST_RESOURCES_DIR) / TEST_DATASET_DIR
        path_dest_dataset = Path(TEST_DATA_DIR) / TEST_DATASET_DIR
        shutil.copytree(str(path_src_dataset.resolve()), str(path_dest_dataset.resolve()))

        path_src_calculators = Path(TEST_RESOURCES_DIR) / TEST_CALCULATORS_DIR
        path_dest_calculators = Path(TEST_DATA_DIR) / TEST_CALCULATORS_DIR
        shutil.copytree(str(path_src_calculators.resolve()), str(path_dest_calculators.resolve()))

        assert not (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()


def test_validate_manifest_json():
    with open(str((Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).resolve())) as json_file:
        data = json.load(json_file)
        assert "dataset" in data
        assert "calculators" in data
        assert data['dataset']['path'] == MATCH_PATH_DATASET
        assert len(data['calculators']) == 2
        assert MATCH_PATH_CALCULATOR_2 in data['calculators'][0]['path']
        assert MATCH_PATH_CALCULATOR_1 in data['calculators'][1]['path']


def test_generate_directories():
    EnvGenerator.generate_directories(TEST_DATA_DIR)
    assert (Path(TEST_DATA_DIR) / TEST_DATASET_DIR).exists()
    assert (Path(TEST_DATA_DIR) / TEST_CALCULATORS_DIR).exists()


def test_generate_new_manifest():
    if (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists() and (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).is_file():
        os.remove(str((Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).resolve()))
    assert not (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()
    with pytest.raises(SystemExit):
        EnvGenerator.generate_manifest(TEST_DATA_DIR, MODE_CREATE_MANIFEST_ONLY)
    assert (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()


def test_generate_existing_manifest():
    EnvGenerator.generate_manifest(TEST_DATA_DIR, MODE_CREATE_MANIFEST_ONLY)
    assert (Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()


prepare_test_env()
test_generate_directories()
test_generate_new_manifest()
test_generate_existing_manifest()
test_validate_manifest_json()
