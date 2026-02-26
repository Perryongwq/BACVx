import os

class directory():

    """
    Instantiate Directories in Project
    """

    def __init__(self):
        self.srcPath = os.path.dirname(os.path.dirname(__file__))
        self.basePath = os.path.dirname(self.srcPath)
        self.acclogPath = os.path.join(self.basePath,'acclog')
        self.accPath = os.path.join(self.basePath,'acc')
        self.dataPath = os.path.join(self.basePath,'data')
        self.prassPath = os.path.join(self.basePath,'prass')
        self.troublePath = os.path.join(self.basePath,'trouble')
