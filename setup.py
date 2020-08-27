from setuptools import setup

APP = ['PBID.py']
DATA_FILES = []
PKGS = [
]

INCLDS = [
    'sys',
    'os',
    'zipfile',
    'json',
    'csv',
    'datetime',
    'tkinter',
]

OPTIONS = {
    #'iconfile': 'icons/PBID-Logo-L',
    'argv_emulation':True,
    'packages': PKGS,
    'includes': INCLDS,
}

setup(
    app = APP,
    data_files = DATA_FILES,
    options = {'py2app': OPTIONS},
    setup_requires = ['py2app'],
)