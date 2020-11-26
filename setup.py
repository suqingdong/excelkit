# -*- encoding: utf8 -*-
import os
import json
import codecs
from setuptools import setup, find_packages


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
version_info = json.load(open(os.path.join(BASE_DIR, 'excelkit', 'version', 'version.json')))


setup(
    name='excelkit',
    version=version_info['version'],
    author=version_info['author'],
    author_email=version_info['author_email'],
    description=version_info['desc'],
    long_description=codecs.open(os.path.join(BASE_DIR, 'README.md'), encoding='utf-8').read(),
    long_description_content_type="text/markdown",
    url='https://github.com/suqingdong/excelkit',
    project_urls={
        'Documentation': 'https://excelkit.readthedocs.io',
        'Tracker': 'https://github.com/suqingdong/excelkit/issues',
    },
    license='BSD License',
    install_requires=codecs.open(os.path.join(BASE_DIR, 'requirements.txt'), encoding='utf-8').read().split('\n'),
    packages=find_packages(),
    include_package_data=True,
    entry_points={'console_scripts': [
        'excelkit = excelkit.bin.main:main',
        'excel-parse = excelkit.bin.main:parse_cli',
        'excel-build = excelkit.bin.main:build_cli',
        'excel-concat = excelkit.bin.main:concat_cli',
    ]},
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Operating System :: OS Independent',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Topic :: Software Development :: Libraries'
    ]
)
