# run pip3 install -e .
# https://godatadriven.com/blog/a-practical-guide-to-using-setup-py/
from setuptools import setup, find_packages
setup(name='report',
      version='0.1.1',
      packages=['report','report.src', 'report.src.operators'],
      install_requires=[
        'numpy',
        'pandas',
        'openpyxl',
        'discord==1.7.3'
    ],)