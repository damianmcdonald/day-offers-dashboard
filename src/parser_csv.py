# parser_csv.py
import csv


def get_records(dataset_csv, num_records, has_header):
    csv_data = []
    with open(dataset_csv) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        record_count = 0
        for row in csv_reader:
            if has_header and record_count == 0:
                record_count += 1
                continue
            json_record = {
                "client": row[0],
                "status": row[1],
                "statusDate": row[2],
                "cloud": row[3],
                "greenfield": row[4],
                "regions": row[5],
                "accounts": row[6],
                "applications": row[7],
                "vpcs": row[8],
                "subnets": row[9],
                "hasConnectivity": row[10],
                "hasPeerings": row[11],
                "hasDirectoryService": row[12],
                "hasAdvancedSecurity": row[13],
                "hasAdvancedLogging": row[14],
                "hasAdvancedMonitoring": row[15],
                "hasAdvancedBackup": row[16],
                "virtualMachines": row[17],
                "buckets": row[18],
                "databases": row[19],
                "hasELB": row[20],
                "hasAutoScripts": row[21],
                "hasOtherServices": row[22],
                "service1": row[23],
                "service2": row[24],
                "service3": row[25],
                "service4": row[26],
                "service5": row[27],
                "phase1EstimatePre": row[28],
                "phase1Estimate": row[29],
                "phase1Deviation": row[30],
                "phase2EstimatePre": row[31],
                "phase2Estimate": row[32],
                "phase2Deviation": row[33],
                "phase3EstimatePre": row[34],
                "phase3Estimate": row[35],
                "phase3Deviation": row[36],
                "phase4EstimatePre": row[37],
                "phase4Estimate": row[38],
                "phase4Deviation": row[39],
                "totalPre": row[40],
                "total": row[41],
                "totalDeviation": row[42],
                "travel": row[43],
                "administered": row[44],
                "geoLocation": row[45],
                "isValid": row[46]
            }
            csv_data.append(json_record)
            record_count += 1
    if num_records > 0:
        # if num_records is greater than zero then we grab a range of the last rows
        return csv_data[-num_records:]
    else:
        # if num_records is 0 then we return a full record set
        return csv_data


def row_count(dataset_csv, has_header):
    line_count = 0
    with open(dataset_csv) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for row in csv_reader:
            line_count += 1
    if has_header:
        # return the row count minus the header row
        return line_count - 1
    else:
        return line_count

