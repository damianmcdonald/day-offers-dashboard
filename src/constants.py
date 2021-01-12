# constants.py
from pathlib import Path


class Constants:

    def __init__(self, path_datadir, mode):
        self.path_datadir = path_datadir
        self.mode = mode

    @staticmethod
    def get_log_dir():
        return "logs"

    @staticmethod
    def get_log_file():
        return "logs/app.log"

    def get_mode(self):
        return self.mode

    def get_path_data_dir(self):
        return Path(self.path_datadir)

    def get_path_data_set(self):
        return Path(self.path_datadir) / "dataset"

    def get_path_calculators(self):
        return Path(self.path_datadir) / "calculators"

    def get_path_manifest_json(self):
        return Path(self.path_datadir) / "manifest.json"