# bangmatching

A script for matching Bang artists and authors. Assumes max 2 artists per author allowed, and that all fic names take the format ficnumber-fictag. 

Takes in Excel spreadsheet (claims Google Form response sheet) with fic names, artist names, and fic rankings, and outputs Excel spreadsheet (matches.xlsx) with fic-artist pairings.

Selects optimal run (optimizing for most first/second/third choice rankings and least fifth choice rankings) out of X runs. Recommend 100-1000 runs.

Runs in, like, O(N^4) time, but shhhh we don't talk about that :)

## Installation
Install dependencies:
```python
pip3 install numpy
```
```python
pip3 install pandas
```

## Usage
```python
python3 matching.py {input spreadsheet filename} {number of fics} {number of runs}
```
