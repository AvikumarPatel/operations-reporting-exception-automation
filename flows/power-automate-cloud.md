# Power Automate Cloud flows

Two cloud flows were created: one for manual testing and one for scheduled operation.

## Manual review flow

`Manually trigger a flow → Run script → Condition`

Condition logic:

```text
int(body('Run_script')?['result']?['exceptionCount']) > 0
```

### True branch
- Send an Outlook email with formatted KPI values.
- Post a Microsoft Teams message in a chat or channel.

### False branch
- Send an Outlook email confirming that no exceptions were detected.

## Scheduled production-style flow

`Recurrence → Run script → Condition`

The same True/False routing is used, but the flow is initiated by a recurrence trigger rather than a manual trigger.

## Office Script

The Excel Online (Business) `Run script` action executes `ReviewOperationsKPIs`, which returns a structured object containing:

- `totalRecords`
- `uniqueOrderIds`
- `totalOrders`
- `totalUnits`
- `overallErrorRate`
- `lateOrderRate`
- `returnRate`
- `ordersPerHour`
- `exceptionCount`
- `status`

The cloud flow uses `exceptionCount` for branching and the other returned values for notification content.
