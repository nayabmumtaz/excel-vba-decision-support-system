Pharmaceutical Capacity Planning Decision Support System (Excel VBA)

Overview
This project implements a decision support system (DSS) built in Microsoft Excel with VBA automation to support pharmaceutical production capacity decisions under demand uncertainty.

The system allows decision-makers to simulate production scenarios for the drug Prizdol, where demand is modelled as a normally distributed random variable. Using simulation and statistical modelling, the DSS evaluates potential capacity levels and provides insights into expected production performance.

The tool integrates Excel modelling, VBA automation, and statistical simulation to support operational decision-making in a pharmaceutical manufacturing environment.

Business Problem

Pharmaceutical manufacturers must determine appropriate production capacity levels while facing uncertain demand.

If capacity is too low:

- unmet demand
- lost revenue
- service level failures

If capacity is too high:
-excess production costs
inefficient resource utilization

This system helps decision-makers evaluate different capacity levels under stochastic demand conditions.

System Architecture
The DSS is organized into five main components:

1. Introduction

Provides an interface describing the system and guiding the user through the model.

2. Inputs

Allows the user to define key model parameters including:
- demand distribution parameters
- service level assumptions
- simulation settings
- production capacity levels

3. Pharmaceutical Model

Core model logic where demand uncertainty and production capacity decisions are evaluated.

4. Simulation Engine

VBA macros automate simulation runs, enabling the model to test multiple capacity scenarios.

5. Outputs & Reporting

Simulation results are summarized and visualized through Excel output sheets and reports.

Key Features
- Excel-based decision support system
- VBA automation for simulation execution
- stochastic demand modelling
- scenario-based capacity analysis
- automated reporting of simulation results

Technologies Used
- Microsoft Excel
- VBA (Visual Basic for Applications)
- Monte Carlo simulation
- Statistical demand modelling

How to Run the Model
1. Download the .xlsm workbook.
2. Open the file in Microsoft Excel.
3. Enable macros.
4. Navigate to the Inputs sheet.
5. Enter model assumptions.
6. Run the simulation macro.
7. Review results in the Outputs and Report sheets.

Example Use Case
Operations managers can use this tool to determine optimal production capacity for a pharmaceutical product under uncertain demand conditions. By testing multiple capacity levels, the model helps identify strategies that balance service levels and operational efficiency.

Author

Nayab Mumtaz
Business Analytics & Data Science (University of Calgary)
