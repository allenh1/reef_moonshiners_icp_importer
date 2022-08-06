# Copyright 2022 Hunter L. Allen
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

from openpyxl import Workbook
from openpyxl import load_workbook

import argparse
import json
import requests
import pandas
import os
import sys

from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QWidget


from html.parser import HTMLParser


class ATIReportParser(HTMLParser):
    def handle_starttag(self, tag, attrs):
        pass

    def handle_endtag(self, tag):
        pass

    def handle_data(self, data):
        if 'var dataTable' in data:
            self.table = json.loads(data.split('\n')[2].replace('data:', '').replace(',\r', ''))


def main():
    _parser = argparse.ArgumentParser(description='Utility to quickly read ATI test results and output an excel sheet')
    _parser.add_argument(
        '--tank-size',
        help='Total water volume (gallons)',
        type=float
    )
    _parser.add_argument(
        '--analysis-id',
        help="Public Analysis ID (number from the url following 'https://lab.atiaquaristik.com/publicAnalysis/'",
        type=str
    )
    args = _parser.parse_args()

    main.analysis_number = str()
    main.tank_size = float()
    if args.analysis_id and args.tank_size:
        main.analysis_number = args.analysis_id
        main.tank_size = args.tank_size
    else:
        # spawn Qt GUI to get information
        app = QApplication([])
        window = QWidget()
        layout = QVBoxLayout()
        button = QPushButton('Get ICP Data')
        # ICP ID Entry
        icp_entry = QHBoxLayout()
        icp_entry.addWidget(QLabel('ATI ICP Analysis ID:'))
        icp_entry_edit = QLineEdit()
        icp_entry.addWidget(icp_entry_edit)
        # Tank Size Entry
        tank_size_entry = QHBoxLayout()
        tank_size_entry.addWidget(QLabel('Tank Size (gallons):'))
        tank_size_edit = QLineEdit()
        tank_size_entry.addWidget(tank_size_edit)
        layout.addLayout(icp_entry)
        layout.addLayout(tank_size_entry)
        layout.addWidget(button)
        window.setLayout(layout)
        # Connect things together
        def on_button_clicked():
            try:
                # Read text fields
                # int(icp_entry_edit.text())
                main.analysis_number = icp_entry_edit.text()
                main.tank_size = float(tank_size_edit.text())
                # close the window
                window.close()
            except ValueError:
                alert = QMessageBox()
                alert.setText('Please verify input information')
                alert.exec()
        button.clicked.connect(on_button_clicked)
        window.show()
        app.exec()

    # open reef moonshiners excel sheet
    workbook = load_workbook(filename='reef_moonshiners.xlsx')

    # grab report
    url = 'https://lab.atiaquaristik.com/publicAnalysis/' + main.analysis_number
    r = requests.get(url)

    parser = ATIReportParser()
    parser.feed(r.text)

    values = dict()

    for idx in range(0, 43):
        values[parser.table[str(idx)]['element']['description_en']] = parser.table[str(idx)]['elements_value']

    # fill out the ICP Assesment Tool
    icp_calc = workbook['ICP Assessment tool']
    icp_calc['F43'] = main.tank_size
    icp_calc['F52'] = values['Salinity']
    icp_calc['F71'] = values['Carbonate hardness']
    icp_calc['F91'] = values['Magnesium']
    icp_calc['F110'] = values['Sulfur']
    icp_calc['F123'] = values['Calcium']
    icp_calc['F142'] = values['Potassium']
    icp_calc['F161'] = values['Bromine']
    icp_calc['F180'] = values['Strontium']
    icp_calc['F199'] = values['Boron']
    icp_calc['F218'] = values['Fluorine']
    icp_calc['F271'] = values['Lithium']
    icp_calc['F284'] = values['Silicon']
    icp_calc['F297'] = values['Iodine']
    icp_calc['F315'] = values['Barium']
    icp_calc['F334'] = values['Molybdenum']
    icp_calc['F353'] = values['Nickel']
    icp_calc['F372'] = values['Manganese']
    icp_calc['F403'] = values['Arsenic']
    icp_calc['F416'] = values['Beryllium']
    icp_calc['F429'] = values['Chrome']
    icp_calc['F460'] = values['Cobalt']
    icp_calc['F491'] = values['Iron']
    icp_calc['F522'] = values['Copper']
    icp_calc['F535'] = values['Selenium']
    icp_calc['F548'] = values['Silver']
    icp_calc['F561'] = values['Vanadium']
    icp_calc['F579'] = values['Zinc']
    icp_calc['F598'] = values['Tin']
    icp_calc['F616'] = values['Aluminium']
    icp_calc['F629'] = values['Lanthanum']

    filename = 'reef_moonshiners_%s.xlsx' % main.analysis_number
    workbook.save(filename=(filename))
    box = QMessageBox()
    box.setText("File saved as '%s'" % (os.path.join(os.getcwd(), filename)))
    box.exec()
    return 0


if __name__ == '__main__':
    sys.exit(main())
