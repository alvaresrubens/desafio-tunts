# Tunts Challenge

**Application to change data in a Google Sheets spreadsheet**

The application developed in the Python programming language (by choice of the candidate), to demonstrate the candidate's programming skills. Was used as a starting point the example of a application in Python, to acess the Google Sheets API, which can be accessed at here [link] (https://developers.google.com/sheets/api/quickstart/python). The full description of the
proposed challenge can be accessed [here] (https://tunts.gupy.io/candidates/applications/57400212/instructions).
The application was developed by Rubens David Alvares da Silva (alvares.rubens@gmail.com).

## Option 1 - Using Docker

### Prerequisites

You need to have the latest version of [Docker] (https://docs.docker.com/docker-for-windows/install/) installed.


### Installation and Execution

Clone or download this repository. To execute, need first to build the docker image:

`` bash
docker build -t desafio .

``

To run after building:

`` bash
docker run -dp 3000:3000 desafio

``

## Option 2 - Using Python

### Prerequisites

You need to have the latest version of [Python] (https://www.python.org/ftp/python/3.9.1/python-3.9.1-amd64.exe) installed.


### Installation and Execution

Clone or download this repository. Install dependencies:
``bash
pip install -r requirements.txt
``

To run this app:

`` bash
python grades.py
``
