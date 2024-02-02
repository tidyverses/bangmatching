# bangmatching

A script for matching Bang artists and authors. Assumes max 2 artists per author allowed, and that all fic identifiers take the format ficnumber-ficdescription (e.g. 05-roleswap, 10-coffee-shop-au, etc). (Fic identifiers are placeholder fic names used to anonymize fics during the claims process, and are the names artists see when looking at the claims spreadsheet.)

Takes in Excel spreadsheet (Google Form response sheet generated during claims process) containing the following fields (additional fields are fine & will be disregarded): 

| Field Name  | Value |
| ------------- | ------------- |
| Preferred Name  | artist's preferred name (string)  |
| First Choice Fic  | identifier of artist's first-choice fic (string, in the format ficnumber-ficname) |
| Second Choice Fic  | identifier of artist's second-choice fic (string, in the format ficnumber-ficname)  |
| Third Choice Fic  | identifier of artist's third-choice fic (string, in the format ficnumber-ficname)  |
| Fourth Choice Fic  | identifier of artist's fourth-choice fic (string, in the format ficnumber-ficname)  |
| Fifth Choice Fic  | identifier of artist's fifth-choice fic (string, in the format ficnumber-ficname)  |

Outputs Excel spreadsheet (matches.xlsx) with fic-artist pairings, selecting optimal run (optimizing for most first/second/third choice rankings and least fifth choice rankings) out of X runs. Recommend 100-1000 runs.

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
python3 matching.py {input spreadsheet filename} {number of fics - aka the highest number associated with a fic} {number of runs}
```
