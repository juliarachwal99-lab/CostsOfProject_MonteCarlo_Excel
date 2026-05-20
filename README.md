# Project Cost Prediction - Monte Carlo Simulation

This repository features an Excel-based **Monte Carlo Simulation** designed to forecast and analyze total project cost probabilities. The model serves as a versatile tool applicable to budgeting and risk assessment for any project type.

### Objective
The primary goal of this simulation is to determine the probability of keeping total investment costs below **PLN 9,100,000 (PLN 9,100K)**. 

### Key Features & Findings
* **Statistical Insights:** Analysis of 5,000 iterations indicates a **48.2% probability** that the project costs will remain under the target budget, with a **51.8% risk** of cost overruns.
* **Data Visualization:** Includes a frequency histogram and a Cumulative Distribution Function (CDF) to identify the most frequent cost outcomes (predominantly within the 9,250–9,500 range).

### Automation (VBA Macro)
The 5,000 random simulation records were dynamically generated using an optimized VBA script. The macro automates the data-gathering process by forcing iterative recalculations of the stochastic model (cell `G10`) and logging the outputs into a structured dataset (range `B23:B5022`). 
