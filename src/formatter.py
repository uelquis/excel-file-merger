
from openpyxl import Workbook
import yaml


class Formatter:
    def __init__(self, config_path):
        with open(config_path, 'r') as f:
            self.config = yaml.safe_load(f)

    def apply(self, workbook: Workbook):
        # TODO: Implement the logic to apply formatting based on the configuration
        pass