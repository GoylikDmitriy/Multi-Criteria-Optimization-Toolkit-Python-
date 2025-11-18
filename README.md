# Multi-Criteria Optimization Toolkit (Python)

Implementation of several classical methods of **multi-criteria optimization** with automation of data loading, normalization, decision-making, and result export.

## Overview

This project provides a Python-based solution for solving **multi-criteria optimization problems** using three widely known decision-making methods:

* **Additive method with weight coefficients**
* **Binary relations method**
* **Main criterion (dominant criterion) method**

The program automates loading input data from Excel, normalizes criteria, determines compromise/consensus points, and computes optimal alternatives for each method.

## Features

- Load criteria matrix from Excel  
- Set optimization directions (max/min)  
- Set weights for criteria (additive method)  
- Data normalization  
- Identify compromise & agreement points  
- Compute results using three optimization methods  
- Export final results to an output file  
- Clean and modular Python implementation  

## Input Data

The program expects an Excel file containing a table of criteria values, for example:

| Alternative | C1 | C2 | C3 | ... |
| ----------- | -- | -- | -- | --- |

**criteria.xlsx** is exmaple of input data file.

## Output

After execution, the program generates an output file containing:

- normalized data  
- values calculated for each method  
- the determined optimal alternatives
