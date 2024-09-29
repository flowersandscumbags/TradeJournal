from setuptools import setup, find_packages

setup(
    name="TradeJournal",
    version="2.0.2",
    author="flowersandscumbags",
    author_email="flowersandscumbags@gmail.com",
    description="A Python script which extracts trading data from PDF files and writes (appends) that data to a multi sheet XLS worksheet.",
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url="https://github.com/flowersandscumbags/TradeJournal",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved ::  GNU General Public License v2 (GPLv2)",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    install_requires=[
        "openpyxl",
        "pdfplumber",
        "tkinter",
        "logging",
        "pandas",
    ],
    entry_points={
        'console_scripts': [
            'tradejournal=copytrades:main',  # Replace 'main' with the actual entry function if different
        ],
    },
)