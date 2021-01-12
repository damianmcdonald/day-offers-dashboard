# test_parser_csv.py
# to get the import statements on the test working, make sure to set PYTHONPATH which includes the src directory
# e.g. export PYTHONPATH=/opt/devel/source/dashboard-flow/src
import parser_csv as CsvParser
import shutil
from pathlib import Path

# tests resources definitions
TEST_RESOURCES_DIR = "resources"
TEST_DATA_DIR = "test_datadir"
TEST_DATASET_DIR = "dataset"
TEST_CALCULATORS_DIR = "calculators"
TEST_MANIFEST_JSON = "manifest.json"
TEST_DATASET_CSV = "static_dataset.csv"
ROW_COUNT_EXPECTED_PRE_WITH_HEADER = 23
ROW_COUNT_EXPECTED_PRE_WITHOUT_HEADER = 24


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


def test_row_count_expected_pre_with_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    assert CsvParser.row_count(str(dataset_csv.resolve()), True) == ROW_COUNT_EXPECTED_PRE_WITH_HEADER


def test_row_count_expected_pre_without_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    assert CsvParser.row_count(str(dataset_csv.resolve()), False) == ROW_COUNT_EXPECTED_PRE_WITHOUT_HEADER


def test_get_all_records_without_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    assert len(CsvParser.get_records(dataset_csv, 0, False)) == ROW_COUNT_EXPECTED_PRE_WITHOUT_HEADER


def test_get_all_records_with_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    assert len(CsvParser.get_records(dataset_csv, 0, True)) == ROW_COUNT_EXPECTED_PRE_WITH_HEADER


def test_get_partial_records_without_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    num_records = 5
    assert len(CsvParser.get_records(dataset_csv, num_records, False)) == num_records


def test_get_partial_records_with_header():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    num_records = 5
    assert len(CsvParser.get_records(dataset_csv, num_records, True)) == num_records


prepare_test_env()
test_row_count_expected_pre_with_header()
test_row_count_expected_pre_without_header()
test_get_all_records_without_header()
test_get_all_records_with_header()
test_get_partial_records_without_header()
test_get_partial_records_with_header()
