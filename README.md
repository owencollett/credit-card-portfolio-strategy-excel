# Credit Card Portfolio Growth & Risk Strategy — Excel

## Overview
This project is a consulting-style Excel case that evaluates a **synthetic portfolio of 3,000 credit-card customers**. The objective is to identify which customers should be targeted for growth while balancing profitability, retention, and credit risk.

The project is intentionally **not a dashboard**. It focuses on the analytical workflow a strategy or consulting analyst would use: clean the data, define business rules, calculate customer economics, segment the portfolio, test scenarios, and make recommendations.

## Business Questions
1. Which customers are attractive candidates for credit-line growth?
2. Which portfolio segments generate the strongest annual contribution?
3. Where is customer retention risk economically meaningful?
4. Which customer acquisition channels produce the strongest average customer economics?
5. How sensitive is a targeted growth program to spend response, incremental losses, and program cost?

## Key Findings
- **956 customers (31.9% of the portfolio)** qualify as Priority Growth candidates.
- Priority Growth customers generate approximately **$325,038** in current annual net contribution and have an average credit score of **720**.
- The base-case line-increase scenario produces approximately **$1,715,582** of incremental annual spend and **$13,856** of incremental net contribution.
- Base-case modeled ROI on program cost is approximately **181.2%**.
- **278 high-priority retention accounts** represent approximately **$25,515** of expected churn-loss exposure.
- **407 high-risk accounts** represent approximately **$56,861** of modeled expected credit loss.
- **Partner** has the highest average annual net contribution by acquisition channel at approximately **$254 per customer**.
- **Platinum** has the highest average annual net contribution by card tier at approximately **$526 per customer**.

## Excel Skills Demonstrated
- XLOOKUP
- SUMIF / SUMIFS / COUNTIF / COUNTIFS
- INDEX / MATCH
- IF / AND / OR logic
- TRIM / PROPER data cleaning
- Customer segmentation and scoring logic
- Financial contribution modeling
- Credit-risk segmentation
- Scenario analysis
- Break-even analysis
- Sensitivity analysis
- Conditional formatting
- Structured Excel tables
- Executive recommendation development

## Workbook Structure
### `README`
Explains the business objective, questions, workflow, and skills demonstrated.

### `Assumptions`
Contains synthetic card economics, risk-loss assumptions, and segmentation thresholds. Keeping assumptions separate makes the model auditable and easy to modify.

### `Raw_Data`
Contains 3,000 synthetic customer records. Some categorical text values intentionally include inconsistent capitalization and whitespace to create a realistic cleaning step.

### `Customer_Analysis`
Cleans the raw data and calculates:
- Risk band
- Annual spend
- Interest revenue
- Interchange revenue
- Rewards cost
- Expected credit loss
- Net annual contribution
- Contribution margin
- Growth score
- Growth segment
- Expected churn loss
- Retention priority

### `Segment_Summary`
Uses formula-driven summary tables to compare:
- Growth segments
- Risk bands
- Card tiers
- Acquisition channels
- Retention-priority groups

### `Scenario_Model`
Tests downside, base, and upside cases for a targeted credit-line increase strategy. The model includes:
- Credit limit increase
- Spend response
- Program cost per account
- Incremental expected loss rate
- Retention lift
- Incremental contribution
- ROI
- Break-even spend response
- Spend-response sensitivity analysis

### `Recommendations`
Translates the analysis into executive actions and links each recommendation to supporting metrics and decision rules.

## Example Decision Logic

**Priority Growth**
- Customer is not High Risk
- Growth Score meets the required threshold
- Net annual contribution is positive

**High Retention Priority**
- Net annual contribution is at least $300
- Customer is not High Risk
- Churn risk is at least 15% **or** satisfaction is 6 or below

## How to Use the Model
1. Open `Credit_Card_Portfolio_Strategy.xlsx`.
2. Review the synthetic assumptions on the `Assumptions` sheet.
3. Trace the formulas in `Customer_Analysis`.
4. Review segment economics in `Segment_Summary`.
5. Change the assumptions in `Scenario_Model` to test alternative strategies.
6. Review the recommendation logic on the `Recommendations` sheet.

## Data
The raw dataset is also included separately at:

`data/customer_portfolio_raw.csv`

## Disclaimer
All customer data, fees, rates, losses, economics, and results in this project are **synthetic and illustrative**. They are not Visa data, issuer data, or real customer data.
