from __future__ import print_function
from constants import Constants
import os.path
import logging
from pathlib import Path
import pygsheets

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

# create a service account for access
# https://medium.com/@denisluiz/python-with-google-sheets-service-account-step-by-step-8f74c26ed28e
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
# https://docs.google.com/spreadsheets/d/16iDNghEdFHDXQJY60EMYhw8cJZRBxBJILYtjSjNIwIU/edit#gid=1152477435
# https://docs.google.com/spreadsheets/d/16iDNghEdFHDXQJY60EMYhw8cJZRBxBJILYtjSjNIwIU/edit#gid=1152477435
WORKSHEET_NAME_DATA = "dataset"
WORKSHEET_NAME_AVG = "averages"
WORKSHEET_NAME_SERVICES = "services"

# https://medium.com/game-of-data/play-with-google-spreadsheets-with-python-301dd4ee36eb

# cell ranges for the DATASET worksheet
RANGES_CELL_FIRST = "A"
RANGES_CELL_LAST = "AU"

# formula builders for the AVERAGES worksheet
FORMULA_AVG_PHASE_1_PRE = f"=AVERAGE({WORKSHEET_NAME_DATA}!AC2:{WORKSHEET_NAME_DATA}!AC"
FORMULA_AVG_PHASE_2_PRE = f"=AVERAGE({WORKSHEET_NAME_DATA}!AF2:{WORKSHEET_NAME_DATA}!AF"
FORMULA_AVG_PHASE_3_PRE = f"=AVERAGE({WORKSHEET_NAME_DATA}!AI2:{WORKSHEET_NAME_DATA}!AI"
FORMULA_AVG_PHASE_4_PRE = f"=AVERAGE({WORKSHEET_NAME_DATA}!AL2:{WORKSHEET_NAME_DATA}!AL"
FORMULA_AVG_TOTAL_PRE = f"=AVERAGE({WORKSHEET_NAME_DATA}!AO2:{WORKSHEET_NAME_DATA}!AO"
FORMULA_AVG_PHASE_1 = f"=AVERAGE({WORKSHEET_NAME_DATA}!AD2:{WORKSHEET_NAME_DATA}!AD"
FORMULA_AVG_PHASE_2 = f"=AVERAGE({WORKSHEET_NAME_DATA}!AG2:{WORKSHEET_NAME_DATA}!AG"
FORMULA_AVG_PHASE_3 = f"=AVERAGE({WORKSHEET_NAME_DATA}!AJ2:{WORKSHEET_NAME_DATA}!AJ"
FORMULA_AVG_PHASE_4 = f"=AVERAGE({WORKSHEET_NAME_DATA}!AM2:{WORKSHEET_NAME_DATA}!AM"
FORMULA_AVG_TOTAL = f"=AVERAGE({WORKSHEET_NAME_DATA}!AP2:{WORKSHEET_NAME_DATA}!AP"
FORMULA_AVG_PHASE_1_DEVIATION = f"=AVERAGE({WORKSHEET_NAME_DATA}!AE2:{WORKSHEET_NAME_DATA}!AE"
FORMULA_AVG_PHASE_2_DEVIATION = f"=AVERAGE({WORKSHEET_NAME_DATA}!AH2:{WORKSHEET_NAME_DATA}!AH"
FORMULA_AVG_PHASE_3_DEVIATION = f"=AVERAGE({WORKSHEET_NAME_DATA}!AK2:{WORKSHEET_NAME_DATA}!AK"
FORMULA_AVG_PHASE_4_DEVIATION = f"=AVERAGE({WORKSHEET_NAME_DATA}!AN2:{WORKSHEET_NAME_DATA}!AN"
FORMULA_AVG_TOTAL_DEVIATION = f"=AVERAGE({WORKSHEET_NAME_DATA}!AQ2:{WORKSHEET_NAME_DATA}!AQ"

# cell references in which to place the formulas for the AVERAGE worksheet
CELL_REF_PHASE_1_PRE = "B2"
CELL_REF_PHASE_2_PRE = "B3"
CELL_REF_PHASE_3_PRE = "B4"
CELL_REF_PHASE_4_PRE = "B5"
CELL_REF_TOTAL_PRE = "B6"
CELL_REF_PHASE_1 = "C2"
CELL_REF_PHASE_2 = "C3"
CELL_REF_PHASE_3 = "C4"
CELL_REF_PHASE_4 = "C5"
CELL_REF_TOTAL = "C6"
CELL_REF_PHASE_1_DEVIATION = "D2"
CELL_REF_PHASE_2_DEVIATION = "D3"
CELL_REF_PHASE_3_DEVIATION = "D4"
CELL_REF_PHASE_4_DEVIATION = "D5"
CELL_REF_TOTAL_DEVIATION = "D6"


def authenticate_gcp_api():
    script_path = os.path.dirname(os.path.realpath("gcp_sheets.py")) + "/src"
    service_file_path = script_path.replace('\\\\', '/') + '/offer-analysis-service-account.json'
    return pygsheets.authorize(service_account_file=service_file_path)


def get_next_worksheet_row(workbook, sheetName):
    cells = workbook.worksheet_by_title(sheetName).get_all_values(include_tailing_empty_rows=False,
                                                                  include_tailing_empty=False,
                                                                  returnas='matrix')
    end_row = len(cells)
    return end_row + 1


def update_worksheet_averages(workbook):
    worksheet_averages = workbook.worksheet_by_title(WORKSHEET_NAME_AVG)

    # find the last row for the DATASET worksheet as we are taking the formula values from that worksheet
    # we need to subtract 1 to get the last active row not the next writable row
    row_limit = get_next_worksheet_row(workbook, WORKSHEET_NAME_DATA) - 1

    cell_formula_avg_phase_1_pre = f"{FORMULA_AVG_PHASE_1_PRE}{row_limit})"
    cell_formula_avg_phase_2_pre = f"{FORMULA_AVG_PHASE_2_PRE}{row_limit})"
    cell_formula_avg_phase_3_pre = f"{FORMULA_AVG_PHASE_3_PRE}{row_limit})"
    cell_formula_avg_phase_4_pre = f"{FORMULA_AVG_PHASE_4_PRE}{row_limit})"
    cell_formula_avg_total_pre = f"{FORMULA_AVG_TOTAL_PRE}{row_limit})"
    cell_formula_avg_phase_1 = f"{FORMULA_AVG_PHASE_1}{row_limit})"
    cell_formula_avg_phase_2 = f"{FORMULA_AVG_PHASE_2}{row_limit})"
    cell_formula_avg_phase_3 = f"{FORMULA_AVG_PHASE_3}{row_limit})"
    cell_formula_avg_phase_4 = f"{FORMULA_AVG_PHASE_4}{row_limit})"
    cell_formula_avg_total = f"{FORMULA_AVG_TOTAL}{row_limit})"
    cell_formula_avg_phase_1_deviation = f"{FORMULA_AVG_PHASE_1_DEVIATION}{row_limit})"
    cell_formula_avg_phase_2_deviation = f"{FORMULA_AVG_PHASE_2_DEVIATION}{row_limit})"
    cell_formula_avg_phase_3_deviation = f"{FORMULA_AVG_PHASE_3_DEVIATION}{row_limit})"
    cell_formula_avg_phase_4_deviation = f"{FORMULA_AVG_PHASE_4_DEVIATION}{row_limit})"
    cell_formula_avg_total_deviation = f"{FORMULA_AVG_TOTAL_DEVIATION}{row_limit})"

    logging.debug(f"End boundary for the AVERAGE formula is cell {row_limit}")
    logging.debug(f"Example formula to be applied: {cell_formula_avg_phase_1_pre}")

    # update pre-estimation averages
    logging.info(f"Updating the pre estimation formulas in worksheet {WORKSHEET_NAME_AVG}")
    worksheet_averages.update_value(CELL_REF_PHASE_1_PRE, cell_formula_avg_phase_1_pre, True)
    worksheet_averages.update_value(CELL_REF_PHASE_2_PRE, cell_formula_avg_phase_2_pre, True)
    worksheet_averages.update_value(CELL_REF_PHASE_3_PRE, cell_formula_avg_phase_3_pre, True)
    worksheet_averages.update_value(CELL_REF_PHASE_4_PRE, cell_formula_avg_phase_4_pre, True)
    worksheet_averages.update_value(CELL_REF_TOTAL_PRE, cell_formula_avg_total_pre, True)

    # update estimation averages
    worksheet_averages.update_value(CELL_REF_PHASE_1, cell_formula_avg_phase_1, True)
    worksheet_averages.update_value(CELL_REF_PHASE_2, cell_formula_avg_phase_2, True)
    worksheet_averages.update_value(CELL_REF_PHASE_3, cell_formula_avg_phase_3, True)
    worksheet_averages.update_value(CELL_REF_PHASE_4, cell_formula_avg_phase_4, True)
    worksheet_averages.update_value(CELL_REF_TOTAL, cell_formula_avg_total, True)

    # update deviation averages
    worksheet_averages.update_value(CELL_REF_PHASE_1_DEVIATION, cell_formula_avg_phase_1_deviation, True)
    worksheet_averages.update_value(CELL_REF_PHASE_2_DEVIATION, cell_formula_avg_phase_2_deviation, True)
    worksheet_averages.update_value(CELL_REF_PHASE_3_DEVIATION, cell_formula_avg_phase_3_deviation, True)
    worksheet_averages.update_value(CELL_REF_PHASE_4_DEVIATION, cell_formula_avg_phase_4_deviation, True)
    worksheet_averages.update_value(CELL_REF_TOTAL_DEVIATION, cell_formula_avg_total_deviation, True)


def update_worksheet_services(data, workbook):
    worksheet_services = workbook.worksheet_by_title(WORKSHEET_NAME_SERVICES)
    row_start = get_next_worksheet_row(workbook, WORKSHEET_NAME_SERVICES) - 1
    services = []
    for record in data:
        if not record['service1'] == "":
            services.append(record['service1'])
        if not record['service2'] == "":
            services.append(record['service2'])
        if not record['service3'] == "":
            services.append(record['service3'])
        if not record['service4'] == "":
            services.append(record['service4'])
        if not record['service5'] == "":
            services.append(record['service5'])

    col_index = 1
    logging.info(f"Writing services to col {col_index} starting at row {row_start}")
    logging.debug(f"Writing services: {str(services)}")
    worksheet_services.update_col(col_index, services, row_offset=row_start)


def update_worksheet_dataset(data, workbook):
    worksheet_data = workbook.worksheet_by_title(WORKSHEET_NAME_DATA)
    row_start = get_next_worksheet_row(workbook, WORKSHEET_NAME_DATA)
    row_end = len(data)
    row_count = 0
    for record in data:
        if row_count <= row_end:
            row_current = row_start + row_count
            range_update = f"{RANGES_CELL_FIRST}{row_current}:{RANGES_CELL_LAST}{row_current}"
            logging.info(f"Writing data for client {record['client']} to sheet {WORKSHEET_NAME_DATA}, range {range_update}")
            record_values = [
                             record['client'],
                             record['status'],
                             record['statusDate'],
                             record['cloud'],
                             record['greenfield'],
                             record['regions'],
                             record['accounts'],
                             record['applications'],
                             record['vpcs'],
                             record['subnets'],
                             record['hasConnectivity'],
                             record['hasPeerings'],
                             record['hasDirectoryService'],
                             record['hasAdvancedSecurity'],
                             record['hasAdvancedLogging'],
                             record['hasAdvancedMonitoring'],
                             record['hasAdvancedBackup'],
                             record['virtualMachines'],
                             record['buckets'],
                             record['databases'],
                             record['hasELB'],
                             record['hasAutoScripts'],
                             record['hasOtherServices'],
                             record['service1'],
                             record['service2'],
                             record['service3'],
                             record['service4'],
                             record['service5'],
                             record['phase1EstimatePre'],
                             record['phase1Estimate'],
                             record['phase1Deviation'],
                             record['phase2EstimatePre'],
                             record['phase2Estimate'],
                             record['phase2Deviation'],
                             record['phase3EstimatePre'],
                             record['phase3Estimate'],
                             record['phase3Deviation'],
                             record['phase4EstimatePre'],
                             record['phase4Estimate'],
                             record['phase4Deviation'],
                             record['totalPre'],
                             record['total'],
                             record['totalDeviation'],
                             record['travel'],
                             record['administered'],
                             record['geoLocation'],
                             record['isValid']
                         ]
            worksheet_data.update_row(row_current, record_values, col_offset=0)
            row_count += 1


def update_workbook(data, sheet_id):
    # grab the handle to the gcp sheets api
    api_handle = authenticate_gcp_api()
    # grab the workbook
    workbook = api_handle.open_by_key(sheet_id)
    # update the worksheets
    update_worksheet_dataset(data, workbook)
    update_worksheet_averages(workbook)
    update_worksheet_services(data, workbook)
