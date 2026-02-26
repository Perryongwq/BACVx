import json
import os
from .directories import directory

class readSettings(directory):

    """
    Instantiate Settings in Project
    """

    def __init__(self,mat="EQA02"):
        super().__init__()
        with open(os.path.join(self.srcPath, 'JSON', 'settings.json')) as f:
            setData = json.load(f)
            self.setData = setData
            self.machine = setData['Machine']
            self.address = setData['Address']
            self.credentials = setData['Credentials']
            self.config = setData['Config']
            self.tolerance = setData['Tolerance']
            self.chipSize = setData['Chip Size']
            self.accuracy = setData['Accuracy']
            self.email = setData.get('Email', {})  # Load email config, default to empty dict if not present

        with open(os.path.join(self.srcPath, 'JSON', self.machine, f'{mat}misc.json')) as g:
            miscData = json.load(g)
            self.miscData = miscData
            self.factor = miscData['Factor']
            self.color = miscData['Color']
            self.font = miscData['Font']
            self.highlight = miscData['Highlight']

        with open(os.path.join(self.srcPath, 'JSON', 'staticnames.json')) as h:
            namesData = json.load(h)
            self.defCode = namesData['Defect Code']
            self.defTape = namesData['Defect Tape']
            self.defSticker = namesData['Defect Sticker']
