# VBA module plan

## Data Export

Question: Save Samples and Results combined or separately?

* **Save Data as XML**

    1. Initial data validation (?)
    1. Create XML document as string
    1. Loop through samples
    1. Loop through results
    1. Request filename/path from user (use sensible default)
    1. Save XML document to file

## Validation

* **Circle Invalid Data**
* **Clear Validation Circles**

## Tools

* **Pre-populate Results table**

    Using "Lab Sample ID" as key, create missing records and populate "Lab Sample ID", "PWS Number", and "Collection Date" columns.
