# RCC Pile Design Template Generator

This repository contains a Python script that generates an Excel workbook for RCC cast-in-situ pile design, including:

- Project input sheet with geotechnical and structural design parameters
- Seismic design sheet with BNBC 2020, ACI, ASTM, and Eurocode considerations
- Pile design checks for axial, flexural, shear, uplift, and serviceability
- Liquefaction and landslide screening inputs
- Dynamic BMD/SFD tab with textbook-style separation and graphs
- Pile sketch tab with a conceptual cross-section and elevation drawing
- Deliverable selector buttons that link to each output sheet

## Requirements

Python 3.9+

Install dependencies:

```bash
pip install -r requirements.txt
```

## Generate the workbook

```bash
python generate_rcc_pile_design_template.py
```

This creates:

```bash
RCC_Pile_Design_Template.xlsx
```

in the project root.

## Workbook structure

1. Cover
2. Inputs
3. Soil_Checks
4. Design_Checks
5. Seismic
6. BMD_SFD
7. Sketch
8. Summary

The workbook is structured to support future use as a standard pile design template, with dynamic seismic inputs and separate output tabs.

## Notes

- This is a design template and should be reviewed by a licensed professional engineer.
- Soil parameters, seismic coefficients, and design actions should be verified against project-specific geotechnical and code requirements.
- The template uses engineering assumptions and should be checked before final design submission.
