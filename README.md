# Amazon Laptop Price Analysis

## Project Overview

This project is an end-to-end data analysis project based on laptop data scraped from Amazon India using Selenium.

The project collects laptop product information, cleans and processes the data, performs exploratory data analysis, creates visualizations, and generates a Word report containing insights from the dataset.

---

## Features

- Web scraping using Selenium
- Automatic pagination handling
- Product name and price extraction
- Data cleaning and preprocessing
- Brand extraction from product titles
- Excel data storage using Pandas
- Data visualization using Matplotlib and Seaborn
- Word report generation using python-docx

---

## Tools Used

- Python
- Selenium
- Pandas
- Matplotlib
- Seaborn
- OpenPyXL
- python-docx

---

## Project Structure and What Each File Does

Amazon_Scraper/
│
├── 01_Code/
│   ├── o1_Scraper.py - Amazon scraper and brand extraction
│   ├── o2_Analysis.py - Generates graphs and performs most of the analysis for the report
│   ├── o3_Report_generator.py - Generates the Word report
│   ├── o4_Data_in_docx_generator.py - Generates Laptop_Data.docx
│
├── 02_Data/
│   ├── Laptop_Data.docx - Laptop data in DOCX format
│   ├── Laptop_Data.pdf - Laptop data in PDF format
│   ├── Laptop_Data.xlsx - Laptop data in Excel format
│
├── 03_Charts/
│   ├── boxplot.png
│   ├── Brand Distribution.png
│   ├── histogram.png
│
├── 04_Report/
│   ├── Price Analysis Report.docx - Final report
