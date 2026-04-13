# Automated-Manual-Fulfillment

This comprehensive suite of Python scripts is tailored for real estate data management, facilitating a streamlined workflow to clean data, identify unique property listings, and merge these listings to meet specific inventory goals. It is structured to work in three main stages, each handled by an individual script:

1. **Main.py (Data Cleaning and Preprocessing):** Cleans Excel files by removing unwanted data and ensuring consistency across entries.
2. **Dupes.py (Duplicate Identification and Extraction):** Identifies unique properties by filtering out duplicates, preparing for the addition of new listings.
3. **Append.py (Data Merging and Goal Fulfillment):** Merges listings from different sources to reach a predefined target number of properties, prioritizing them based on potential value.

## Detailed Workflow

### Main.py: Data Cleaning and Preprocessing

- **Objective:** Initiates the process by cleaning the input data files, ensuring that they are free from prohibited tags, unwanted names, and duplicate entries.
- **Key Actions:** Removes specific columns, fills missing values, filters out rows based on predefined criteria, and sorts data for optimal processing.
- **Outcome:** Produces cleaned Excel files ready for further processing, laying the groundwork for identifying unique property listings.

### Dupes.py: Duplicate Identification and Extraction

- **Objective:** Builds upon the cleaned data by identifying and extracting unique property listings from a dataset potentially containing duplicates.
- **Key Actions:** Aggregates duplicate entries across multiple sources into a single database, then isolates unique listings not found in this database.
- **Outcome:** Yields a set of unique property listings (extras) that are not present in the initial dataset, ensuring the uniqueness of property data.

### Append.py: Data Merging and Goal Fulfillment

- **Objective:** Completes the workflow by merging the unique property listings identified in the previous step with the main dataset to meet a specified total goal.
- **Key Actions:** Dynamically calculates the number of properties needed to reach the goal, selectively merging listings based on their assessed value.
- **Outcome:** Generates a final, consolidated Excel file that meets the predefined property count goal, prioritized by the listings' potential value.

## Setup and Requirements

Before running the scripts, ensure your Python environment is set up with Python 3.x and the necessary libraries (`pandas` and `openpyxl`). Organize your Excel files according to the input requirements of each script, and adjust the scripts' parameters to match your dataset and goals.

## Execution Guide

1. **Run main.py:** Start with the data cleaning and preprocessing script to prepare your data files. Ensure all unwanted data is filtered out and the data is consistent.
2. **Proceed with dupes.py:** After cleaning, identify and extract unique property listings from the dataset, preparing a list of additional properties to potentially include.
3. **Finalize with append.py:** Merge the newly identified listings with the main dataset to ensure the total number of properties meets your operational goal, with a focus on adding the most valuable properties first.

### Important Notes

- Each script is designed to operate sequentially but can be adjusted for standalone use if required.
- Customize folder paths, criteria for filtering and sorting, and the total goal in the scripts to fit your specific needs.

This suite is an invaluable tool for real estate professionals looking to optimize their property data management processes, from initial cleaning to the strategic merging of listings to meet inventory objectives.


