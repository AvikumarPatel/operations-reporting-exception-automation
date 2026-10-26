# Test results

All values below come from the included fictional practice dataset and are for portfolio demonstration only.

## Desktop processing test

- Records processed: **600**
- Unique order IDs: **590**
- Total orders processed: **70,009**
- Total units picked: **331,723**
- Overall error rate: **1.12%**
- Late order rate: **2.50%**
- Return rate: **1.46%**
- Overall orders/hour: **15.17**
- Exception records: **381**

Observed exception/data-quality flags included:

- `HIGH_ERROR_RATE`
- `HIGH_LATE_ORDER_RATE`
- `HIGH_RETURN_RATE`
- `LOW_PRODUCTIVITY`
- `MISSING_EMPLOYEE_ID`
- `MISSING_WAREHOUSE`
- `DUPLICATE_ORDER_ID`

## Power Automate Desktop test

Confirmed sequence:

`Launch Excel → Run VBA → Read KPI cells → Evaluate condition → Display alert → Save → Close Excel`

Excel closed automatically after the flow completed.

## Office Script test

`ReviewOperationsKPIs` ran successfully and wrote:

- `Cloud Review Status = Review Required`
- `Last Cloud Review = timestamp`

It also returned a structured KPI object for Power Automate Cloud.

## Cloud True-branch test

With `exceptionCount > 0`:

- Condition evaluated True.
- Outlook review-required email was received.
- Email included formatted KPI values and status.
- Teams notification action was configured in the True branch.

## Cloud False-branch test

For branch validation, `KPI_Summary!B12` was temporarily set to `0`:

- Condition evaluated False.
- Outlook no-exceptions email was received.
- The workbook value was restored after testing.

## Scheduled flow

A recurrence-based production-style version was configured with the same Office Script and True/False routing.
