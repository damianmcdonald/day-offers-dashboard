# test_main.py
# to get the import statements on the test working, make sure to set PYTHONPATH which includes the src directory
# e.g. export PYTHONPATH=/opt/devel/source/dashboard-flow/src
import main
import sys
import shutil
from pathlib import Path
import json
import pytest

# tests resources definitions
MODE_CREATE_MANIFEST_ONLY = "CREATE_MANIFEST_ONLY"
MODE_BUILD_DATASET_ONLY = "BUILD_DATASET_ONLY"
MODE_PUBLISH_DATASET_ONLY = "PUBLISH_DATASET_ONLY"
MODE_CREATE_BUILD_PUBLISH = "CREATE_BUILD_PUBLISH"
MODES = [MODE_CREATE_MANIFEST_ONLY, MODE_BUILD_DATASET_ONLY, MODE_PUBLISH_DATASET_ONLY, MODE_CREATE_BUILD_PUBLISH]
WORKBOOK_ID = '1U82w5zztj_OKOmRaOwcqqGq-8EFiiLJ8bPwXniGTPb0'
MODE_INVALID = "INVALID_MODE"
PROGRAM_NAME = "day-offers-dashboard"
TEST_RESOURCES_DIR = "resources"
TEST_DATA_DIR = "test_datadir"
TEST_DATASET_DIR = "dataset"
TEST_CALCULATORS_DIR = "calculators"
TEST_MANIFEST_JSON = "manifest.json"
TEST_DATA_DIR_BAD_PATH = "bad_path"
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
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_CREATE_MANIFEST_ONLY]
    with pytest.raises(SystemExit):
        main.main()
        with open(str((Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).resolve())) as json_file:
            data = json.load(json_file)
            assert "dataset" in data
            assert "calculators" in data
            assert data['dataset']['path'] == MATCH_PATH_DATASET
            assert len(data['calculators']) == 2
            assert MATCH_PATH_CALCULATOR_2 in data['calculators'][0]['path']
            assert MATCH_PATH_CALCULATOR_1 in data['calculators'][1]['path']


def test_main_valid_create_manifest_only():
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_CREATE_MANIFEST_ONLY, '-sid', WORKBOOK_ID]
    with pytest.raises(SystemExit):
        main.main()
        assert Path(Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()


def test_main_valid_build_dataset_only():
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_BUILD_DATASET_ONLY, '-sid', WORKBOOK_ID]
    with pytest.raises(SystemExit):
        main.main()
        assert Path(Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()
        assert Path(MATCH_PATH_DATASET).exists()


def test_main_valid_publish_dataset_only():
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_PUBLISH_DATASET_ONLY, '-sid', WORKBOOK_ID]
    with pytest.raises(SystemExit):
        main.main()
        assert Path(Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()
        assert Path(MATCH_PATH_DATASET).exists()


def test_main_valid_create_build_publish():
    prepare_test_env()
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_CREATE_BUILD_PUBLISH, '-sid', WORKBOOK_ID]
    main.main()
    assert Path(Path(TEST_DATA_DIR) / TEST_MANIFEST_JSON).exists()
    assert Path(MATCH_PATH_DATASET).exists()


def test_main_invalid_no_args():
    with pytest.raises(Exception):
        assert main.main()


def test_main_invalid_mode():
    sys.argv = ["main.py", '-d', str(Path(TEST_DATA_DIR).resolve()), '-m', MODE_INVALID, '-sid', WORKBOOK_ID]
    with pytest.raises(SystemExit):
        assert main.main()


def test_main_bad_path():
    with pytest.raises(SystemExit):
        sys.argv = ["main.py", '-d', TEST_DATA_DIR_BAD_PATH, '-m', MODE_BUILD_DATASET_ONLY, '-sid', WORKBOOK_ID]
        main.main()


def test_main_missing_sheet_id():
    with pytest.raises(SystemExit):
        sys.argv = ["main.py", '-d', TEST_DATA_DIR_BAD_PATH, '-m', MODE_BUILD_DATASET_ONLY, '-sid', ""]
        main.main()


prepare_test_env()
test_main_valid_create_manifest_only()
test_main_valid_build_dataset_only()
test_main_valid_publish_dataset_only()
test_main_valid_create_build_publish()
test_main_invalid_no_args()
test_main_invalid_mode()
test_main_missing_sheet_id()
test_validate_manifest_json()
