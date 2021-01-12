import gcp_sheets
import parser_csv as CsvParser
import shutil
from pathlib import Path

# tests resources definitions
WORKBOOK_ID = '1U82w5zztj_OKOmRaOwcqqGq-8EFiiLJ8bPwXniGTPb0'
TEST_RESOURCES_DIR = "resources"
TEST_DATA_DIR = "test_datadir"
TEST_DATASET_DIR = "dataset"
TEST_CALCULATORS_DIR = "calculators"
TEST_MANIFEST_JSON = "manifest.json"
TEST_DATASET_CSV = "static_dataset.csv"


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


def test_update_workbook():
    dataset_csv = Path(TEST_DATA_DIR) / TEST_DATASET_DIR / TEST_DATASET_CSV
    num_records = 2
    data = CsvParser.get_records(dataset_csv, num_records, True)
    gcp_sheets.update_workbook(data, WORKBOOK_ID)


prepare_test_env()
test_update_workbook()

