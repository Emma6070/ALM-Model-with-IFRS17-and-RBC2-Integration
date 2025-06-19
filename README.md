# Insurance ALM Model with IFRS17 and RBC2 Integration

## Description
An integrated Asset-Liability Management (ALM) model that combines Prophet forecasting with IFRS17 and RBC2 regulatory requirements for insurance companies. The model provides analysis including technical provisions calculation, capital adequacy assessment, and stress testing capabilities.

## Features
- IFRS17 Technical Provisions calculation
  - Best Estimate Liability (BEL)
  - Risk Adjustment
  - Contractual Service Margin (CSM)
- RBC2 Capital Requirements assessment
  - C1-C4 risk components
  - Capital adequacy ratio
  - Available capital calculation
- Prophet-based asset value forecasting
- Stress testing scenarios
- Interactive dashboard
- Automated Excel reporting
- Data visualization

## Requirements
```bash
pip install -r requirements.txt
```

Required packages:
- prophet>=1.1.4
- pandas>=2.0.0
- numpy>=1.24.0
- matplotlib>=3.7.0
- openpyxl>=3.1.0

## Installation
1. Clone the repository
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage
1. Update the historical data in the script:
```python
historical_data = {
    'ds': pd.date_range(start='2023-01-01', periods=12, freq='M'),
    'y': [1000000, 1020000, 1040000, ...]  # Your asset values
}
```

2. Adjust model parameters in the Assumptions section if needed
3. Run the script:
```bash
python prophetRBC2IFRS17.py
```

## Output
- Excel workbook (`ALM_RBC2_IFRS17_Model.xlsx`) containing:
  - IFRS17 Technical Provisions
  - RBC2 Capital Requirements
  - Asset Projections
  - Risk Analysis
  - Stress Testing Results
  - Performance Dashboard
- Visualization plots for key metrics


## Notes
- Customize risk factors and parameters according to your jurisdiction
- Adjust stress testing scenarios based on your risk appetite
- Model assumptions should be reviewed periodically
